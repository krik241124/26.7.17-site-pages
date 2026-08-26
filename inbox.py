# -*- coding: utf-8 -*-
"""
ArkSwift Outlook Inbox Triage MVP

放置位置：C:\\code\\outreach\\inbox.py
输出位置：C:\\code\\outreach\\output\\outlook_triage.xlsx

功能：
- 读取 Classic Outlook Inbox 全部历史邮件（包括已读）
- 读取 Sent Items 建立已发送会话索引
- 分类：发送失败 / 自动回复 / 真人回复 / 其他
- 对真人回复标记：高合作意向候选 / 明确拒绝候选 / 待人工判断
- 对退信尽量提取失败收件邮箱
- 只读 Outlook：不删除、不移动、不标已读、不自动回复

依赖：
    pip install openpyxl pywin32
"""

import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import pythoncom
import win32com.client
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
REPORT_PATH = OUTPUT_DIR / "outlook_triage.xlsx"

# 多账号时可填写企业邮箱；留空使用 Outlook 默认账号。
MAILBOX_ACCOUNT = ""
SCAN_INBOX_SUBFOLDERS = True
BODY_PREVIEW_CHARS = 2500
MAX_INBOX_ITEMS = None  # None = 全量历史邮件

OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", re.I)
REPLY_PREFIX_RE = re.compile(r"^\s*((re|fw|fwd|aw|sv|答复|回复|转发)\s*:\s*)+", re.I)

BOUNCE_SUBJECT_PATTERNS = [
    r"delivery status notification", r"delivery failure", r"delivery failed",
    r"undeliverable", r"returned mail", r"failure notice", r"mail delivery failed",
    r"message blocked", r"couldn['’]t be delivered", r"could not be delivered",
    r"non[- ]delivery",
]
BOUNCE_BODY_PATTERNS = [
    r"couldn['’]t be delivered", r"could not be delivered", r"wasn['’]t found at",
    r"was not found at", r"delivery status notification", r"delivery has failed",
    r"delivery failed", r"undeliverable", r"unknown to address",
    r"recipient address rejected", r"user unknown", r"mailbox unavailable",
    r"message blocked", r"delivery loop", r"5\.1\.1", r"5\.1\.8",
    r"permanent failure",
]
BOUNCE_SENDER_HINTS = ["mailer-daemon", "mail delivery subsystem", "postmaster", "microsoft outlook", "microsoft exchange"]

AUTO_SUBJECT_PATTERNS = [
    r"automatic reply", r"auto(?:matic)?[- ]?reply", r"out of office", r"\booo\b",
    r"away from (?:the )?office", r"request received", r"ticket received",
    r"case received", r"support request received",
]
AUTO_BODY_PATTERNS = [
    r"this is an automatic(?:ally generated)? (?:reply|response|message)",
    r"automated (?:reply|response|message)",
    r"we have received your (?:request|message|email)",
    r"we(?:'|’)ve received your (?:request|message|email)",
    r"your (?:request|ticket|case).{0,80}(?:has been|was) received",
    r"we will get back to you", r"we(?:'|’)ll get back to you",
    r"we will respond within", r"we(?:'|’)ll respond within",
    r"ticket (?:number|#|id)", r"support ticket",
    r"solicitarea dumneavoastr.{0,80}a fost primit",
    r"urmeaz.{0,40}s.{0,20}fie examinat", r"vom reveni cu un r.{0,20}spuns",
]

SYSTEM_OTHER_SUBJECT_PATTERNS = [
    r"user activation", r"account activation", r"activate your account",
    r"verify your email", r"email verification", r"password reset",
    r"security alert", r"login code", r"one[- ]time pass", r"newsletter",
    r"weekly digest", r"daily digest",
]
SYSTEM_OTHER_BODY_PATTERNS = [
    r"new .* account has been created for you", r"click .* to activate your account",
    r"select a password", r"verify your email", r"unsubscribe", r"manage your preferences",
]

POSITIVE_INTENT_PATTERNS = [
    r"\bopen to\b", r"\binterested\b", r"\bcollaborat", r"\bpartnership",
    r"\blink exchange", r"\bguest (?:post|contribution)", r"\bsponsored",
    r"\bpaid\b", r"\bunpaid\b", r"\bmedia kit\b", r"\beditorial guidelines",
    r"\brequirements?\b", r"\bcriteria\b", r"\brate(?:s)?\b", r"\bpricing\b",
    r"\bhappy to\b", r"\blet me know\b", r"\bwe accept\b", r"\bcontent exchange\b",
]
NEGATIVE_INTENT_PATTERNS = [
    r"\bnot interested\b", r"\bnot open to\b", r"\bwe do not accept\b",
    r"\bwe don['’]t accept\b", r"\bwe cannot collaborate\b",
    r"\bwe can['’]t collaborate\b", r"\bunable to collaborate\b",
    r"\bno partnerships?\b", r"\bnot accepting\b", r"\bplease remove\b",
    r"\bdo not contact\b", r"\bno thank",
]


def safe_get(obj, attr, default=""):
    try:
        value = getattr(obj, attr)
        return default if value is None else value
    except Exception:
        return default


def clean_text(value):
    return "" if value is None else str(value).replace("\x00", "").strip()


def normalize_subject(subject):
    subject = REPLY_PREFIX_RE.sub("", clean_text(subject))
    return re.sub(r"\s+", " ", subject).strip().lower()


def regex_any(patterns, text):
    text = clean_text(text)
    return any(re.search(p, text, re.I | re.S) for p in patterns)


def extract_emails(text):
    values = EMAIL_RE.findall(str(text or "").replace(r"\@", "@"))
    out, seen = [], set()
    for value in values:
        email = value.lower().strip(" <>[](){}.,;:")
        if email and email not in seen:
            seen.add(email)
            out.append(email)
    return out


def get_sender_email(item):
    addr = clean_text(safe_get(item, "SenderEmailAddress"))
    if "@" in addr:
        return addr.lower()
    try:
        if clean_text(item.SenderEmailType).upper() == "EX":
            ex = item.Sender.GetExchangeUser()
            if ex:
                smtp = clean_text(ex.PrimarySmtpAddress)
                if smtp:
                    return smtp.lower()
    except Exception:
        pass
    return addr.lower()


def fmt_time(value):
    try:
        return value.strftime("%Y-%m-%d %H:%M:%S") if value else ""
    except Exception:
        return str(value or "")


def get_account_and_store(namespace):
    wanted = MAILBOX_ACCOUNT.strip().lower()
    if wanted:
        for account in namespace.Accounts:
            try:
                smtp = clean_text(account.SmtpAddress).lower()
            except Exception:
                smtp = ""
            if smtp == wanted:
                return account, account.DeliveryStore
        raise RuntimeError(f"Outlook 中找不到账号：{MAILBOX_ACCOUNT}")
    inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    try:
        account = namespace.Accounts.Item(1)
    except Exception:
        account = None
    return account, inbox.Store


def walk_folder(folder):
    yield folder
    if not SCAN_INBOX_SUBFOLDERS:
        return
    try:
        subs = folder.Folders
        for i in range(1, subs.Count + 1):
            yield from walk_folder(subs.Item(i))
    except Exception:
        return


def build_sent_index(sent_folder):
    conv_ids, subjects = set(), set()
    print("正在建立 Sent Items 会话索引...")
    items = sent_folder.Items
    try:
        items.Sort("[SentOn]", True)
    except Exception:
        pass
    total = items.Count
    for i in range(1, total + 1):
        try:
            item = items.Item(i)
        except Exception:
            continue
        mc = clean_text(safe_get(item, "MessageClass"))
        if not (mc.startswith("IPM.Note") or mc.startswith("REPORT.")):
            continue
        cid = clean_text(safe_get(item, "ConversationID"))
        topic = normalize_subject(safe_get(item, "ConversationTopic"))
        subj = normalize_subject(safe_get(item, "Subject"))
        if cid:
            conv_ids.add(cid)
        if topic:
            subjects.add(topic)
        if subj:
            subjects.add(subj)
        if i % 500 == 0:
            print(f"  已索引 {i}/{total}")
    print(f"Sent Items：{len(conv_ids)} 个会话，{len(subjects)} 个主题。\n")
    return conv_ids, subjects


def extract_failed_recipient(subject, body, own_email=""):
    text = f"{subject}\n{body}".replace(r"\@", "@")
    patterns = [
        r"your message to\s+([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
        r"message to\s+([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
        r"recipient(?: address)?[:\s]+([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
        r"delivery to\s+([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
    ]
    for p in patterns:
        m = re.search(p, text, re.I | re.S)
        if m:
            return m.group(1).lower()
    own_email = own_email.lower().strip()
    for email in extract_emails(text):
        if own_email and email == own_email:
            continue
        if any(x in email for x in ("mailer-daemon", "postmaster", "microsoftexchange")):
            continue
        return email
    return ""


def classify(item, sent_conv_ids, sent_subjects, own_email):
    subject = clean_text(safe_get(item, "Subject"))
    body = clean_text(safe_get(item, "Body"))
    sender_name = clean_text(safe_get(item, "SenderName"))
    sender_email = get_sender_email(item)
    mc = clean_text(safe_get(item, "MessageClass"))
    cid = clean_text(safe_get(item, "ConversationID"))
    topic = clean_text(safe_get(item, "ConversationTopic"))
    combined = f"{subject}\n{sender_name}\n{sender_email}\n{body}"

    is_ndr = mc.upper().startswith("REPORT.") and "NDR" in mc.upper()
    sender_bounce = any(h in f"{sender_name} {sender_email}".lower() for h in BOUNCE_SENDER_HINTS)
    if is_ndr or regex_any(BOUNCE_SUBJECT_PATTERNS, subject) or (sender_bounce and regex_any(BOUNCE_BODY_PATTERNS, body)):
        return "发送失败", "需要重新找邮箱", "高", "检测到 NDR/退信特征。", extract_failed_recipient(subject, body, own_email)

    if regex_any(AUTO_SUBJECT_PATTERNS, subject) or regex_any(AUTO_BODY_PATTERNS, body):
        return "自动回复", "无需处理", "中高", "检测到自动回复/工单确认/Out of Office 模板。", ""

    if regex_any(SYSTEM_OTHER_SUBJECT_PATTERNS, subject) or regex_any(SYSTEM_OTHER_BODY_PATTERNS, body):
        return "其他", "系统/其他", "高", "检测到账户激活、验证、Newsletter 等系统邮件。", ""

    norm_subject = normalize_subject(subject)
    norm_topic = normalize_subject(topic)
    matched_conv = bool(cid and cid in sent_conv_ids)
    matched_subject = bool((norm_subject and norm_subject in sent_subjects) or (norm_topic and norm_topic in sent_subjects))
    reply_prefix = bool(REPLY_PREFIX_RE.match(subject or ""))

    if matched_conv or matched_subject or reply_prefix:
        pos = sum(bool(re.search(p, combined, re.I | re.S)) for p in POSITIVE_INTENT_PATTERNS)
        neg = sum(bool(re.search(p, combined, re.I | re.S)) for p in NEGATIVE_INTENT_PATTERNS)
        if neg:
            intent = "明确拒绝候选"
        elif pos >= 2:
            intent = "高合作意向候选"
        else:
            intent = "真人回复-待判断"
        if matched_conv:
            return "真人回复", intent, "高", "ConversationID 与已发送邮件匹配，且未命中系统规则。", ""
        if matched_subject:
            return "真人回复", intent, "中高", "主题与已发送邮件匹配，且未命中系统规则。", ""
        return "真人回复", intent, "中", "带 Re/AW/SV 等回复前缀，且未命中系统规则。", ""

    return "其他", "其他/待人工抽查", "中", "没有足够证据证明是外联真人回复。", ""


HEADERS = [
    "分类", "意向/动作", "置信度", "收到时间", "未读", "发件人姓名", "发件人邮箱",
    "主题", "失败收件邮箱", "匹配原因", "正文预览", "ConversationID",
    "ConversationTopic", "MessageClass", "Outlook EntryID", "所在文件夹",
]


def new_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name in ["汇总", "全部", "发送失败", "自动回复", "真人回复", "其他"]:
        wb.create_sheet(name)
    for name in ["全部", "发送失败", "自动回复", "真人回复", "其他"]:
        ws = wb[name]
        ws.append(HEADERS)
        for c in ws[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:P1"
    return wb


def style_sheet(ws):
    widths = {1:12,2:20,3:10,4:20,5:8,6:24,7:32,8:55,9:32,10:55,11:90,12:32,13:45,14:28,15:36,16:28}
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = width
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)


def write_summary(wb, counts, unread, total):
    ws = wb["汇总"]
    ws["A1"] = "ArkSwift Outlook 邮件整理 MVP"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A3"], ws["B3"] = "生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws["A4"], ws["B4"] = "扫描邮件总数", total
    ws["A5"], ws["B5"] = "其中未读", unread
    ws.append([])
    ws.append(["分类", "数量", "建议动作"])
    for c in ws[7]:
        c.font = Font(bold=True)
    ws.append(["发送失败", counts["发送失败"], "集中重新寻找正确邮箱"])
    ws.append(["自动回复", counts["自动回复"], "无需阅读"])
    ws.append(["真人回复", counts["真人回复"], "优先人工阅读并推进合作"])
    ws.append(["其他", counts["其他"], "系统邮件/边界案例，抽查"])
    ws["A14"] = "MVP 原则"
    ws["A14"].font = Font(bold=True)
    ws["A15"] = "只读 Outlook；不删除、不移动、不标已读、不自动回复。分类偏保守，宁可多留边界案例，也避免漏掉真人回复。"
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 48


def main():
    if sys.platform != "win32":
        raise RuntimeError("该脚本需要 Windows + Classic Outlook。")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print("=" * 72)
    print("ArkSwift Outlook Inbox Triage MVP")
    print("=" * 72)
    print("只读模式：不会修改任何 Outlook 邮件")
    print(f"输出：{REPORT_PATH}\n")

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        account, store = get_account_and_store(ns)
        try:
            own_email = clean_text(account.SmtpAddress).lower() if account else ""
        except Exception:
            own_email = ""
        print(f"邮箱：{own_email or 'Outlook 默认账号'}")

        inbox = store.GetDefaultFolder(OL_FOLDER_INBOX)
        sent = store.GetDefaultFolder(OL_FOLDER_SENT_MAIL)
        sent_conv_ids, sent_subjects = build_sent_index(sent)

        wb = new_workbook()
        counts = {"发送失败": 0, "自动回复": 0, "真人回复": 0, "其他": 0}
        total = 0
        unread_count = 0

        folders = list(walk_folder(inbox))
        print(f"准备扫描 Inbox 文件夹：{len(folders)} 个\n")

        stop = False
        for folder in folders:
            if stop:
                break
            folder_name = clean_text(safe_get(folder, "FolderPath")) or clean_text(safe_get(folder, "Name"))
            print(f"扫描：{folder_name}")
            items = folder.Items
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass
            for i in range(1, items.Count + 1):
                if MAX_INBOX_ITEMS is not None and total >= MAX_INBOX_ITEMS:
                    stop = True
                    break
                try:
                    item = items.Item(i)
                except Exception:
                    continue
                mc = clean_text(safe_get(item, "MessageClass"))
                if not (mc.startswith("IPM.Note") or mc.startswith("REPORT.")):
                    continue

                category, intent, confidence, reason, failed_recipient = classify(item, sent_conv_ids, sent_subjects, own_email)
                unread = bool(safe_get(item, "UnRead", False))
                if unread:
                    unread_count += 1
                body = clean_text(safe_get(item, "Body"))
                row = [
                    category, intent, confidence,
                    fmt_time(safe_get(item, "ReceivedTime", None)),
                    "是" if unread else "否",
                    clean_text(safe_get(item, "SenderName")),
                    get_sender_email(item),
                    clean_text(safe_get(item, "Subject")),
                    failed_recipient,
                    reason,
                    body[:BODY_PREVIEW_CHARS],
                    clean_text(safe_get(item, "ConversationID")),
                    clean_text(safe_get(item, "ConversationTopic")),
                    mc,
                    clean_text(safe_get(item, "EntryID")),
                    folder_name,
                ]
                wb["全部"].append(row)
                wb[category].append(row)
                counts[category] += 1
                total += 1
                if total % 100 == 0:
                    print(f"  已扫描 {total} | 失败 {counts['发送失败']} | 自动 {counts['自动回复']} | 真人 {counts['真人回复']} | 其他 {counts['其他']}")

        write_summary(wb, counts, unread_count, total)
        for name in ["全部", "发送失败", "自动回复", "真人回复", "其他"]:
            style_sheet(wb[name])
        try:
            wb.save(REPORT_PATH)
        except PermissionError:
            raise RuntimeError(f"无法写入 {REPORT_PATH}，请先关闭这个 Excel。")

        print("\n" + "=" * 72)
        print("整理完成")
        print("=" * 72)
        print(f"共扫描：{total}")
        print(f"其中未读：{unread_count}")
        print(f"发送失败：{counts['发送失败']}")
        print(f"自动回复：{counts['自动回复']}")
        print(f"真人回复：{counts['真人回复']}")
        print(f"其他：{counts['其他']}")
        print(f"\n报告：{REPORT_PATH}")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n用户停止。Outlook 邮件没有被修改。")
    except Exception as exc:
        print("\n脚本停止：")
        print(exc)
        sys.exit(1)
