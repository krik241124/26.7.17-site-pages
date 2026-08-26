# -*- coding: utf-8 -*-
"""
ArkSwift Outlook 批量外联发送脚本
文件名：send.py

使用方式：
1. 将本脚本与 outreach_emails.xlsx 放在同一目录。
2. 安装依赖：
       pip install openpyxl pywin32
3. 确保 Windows 上已安装并登录 Classic Outlook（经典 Outlook）。
4. 关闭正在打开的 outreach_emails.xlsx，避免 Excel 锁文件。
5. 运行：
       python send.py
6. 输入 SEND 后开始正式发送。

功能：
- 自动识别 Domain / To / Subject / Body 列
- 邮箱单元格为空：跳过，并写“无邮箱”
- 单元格有多个邮箱：用正则直接提取，不依赖 ; / ； / , 等分隔符
- 每个邮箱单独发送一封，避免多个联系人互相看到
- 每天最多发送 100 封（按实际收件邮箱计数）
- 每封随机等待 25~45 秒
- 每成功发送一封立即保存 Excel
- 支持断点续跑，不重复发送同一行中已经成功发送的邮箱
- 自动新增：发送状态 / 已发送邮箱 / 失败邮箱 / 发送记录 / 发送备注
"""

import html
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import pythoncom
import win32com.client


# ============================================================
# 基本配置
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
XLSX_PATH = BASE_DIR / "output" / "outreach_emails.xlsx"

DAILY_LIMIT = 100

# 每封邮件之间随机等待，单位：秒
MIN_DELAY_SECONDS = 25
MAX_DELAY_SECONDS = 45

# 每发送多少封后额外休息一次
BATCH_SIZE = 20

# 额外休息时间，单位：秒
BATCH_PAUSE_MIN_SECONDS = 90
BATCH_PAUSE_MAX_SECONDS = 180

# 如果 Outlook 中只有一个发件账号，保持为空即可。
# 如果 Outlook 同时登录多个邮箱，可以填企业邮箱，例如：
# SENDER_ACCOUNT = "jandy.jang@arkswift.com"
SENDER_ACCOUNT = ""

# 正式发送前要求在终端输入 SEND
REQUIRE_CONFIRMATION = True


# ============================================================
# 表头别名
# ============================================================

HEADER_ALIASES = {
    "domain": [
        "Domain (域名)",
        "Domain",
        "域名",
    ],
    "to": [
        "To (收件邮箱)",
        "To",
        "收件邮箱",
        "Email",
        "邮箱",
        "邮箱地址",
    ],
    "subject": [
        "Subject (邮件主题)",
        "Subject",
        "邮件主题",
        "主题",
    ],
    "body": [
        "Body (邮件正文)",
        "Body",
        "邮件正文",
        "正文",
    ],
}

OUTPUT_HEADERS = [
    "发送状态",
    "已发送邮箱",
    "失败邮箱",
    "发送记录",
    "发送备注",
]


# ============================================================
# 工具函数
# ============================================================

EMAIL_RE = re.compile(
    r"(?<![A-Za-z0-9._%+\-])"
    r"[A-Za-z0-9._%+\-]+"
    r"@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"
    r"(?![A-Za-z0-9._%+\-])",
    re.IGNORECASE,
)


def normalize_header(value):
    """标准化表头，用于兼容中英文列名。"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).strip().lower()


def find_column(ws, aliases):
    """根据多个可能的表头名称寻找列号。"""
    alias_set = {normalize_header(x) for x in aliases}

    for cell in ws[1]:
        if normalize_header(cell.value) in alias_set:
            return cell.column

    return None


def ensure_column(ws, header_name):
    """
    找到指定输出列；不存在则追加到最右侧。
    返回列号。
    """
    target = normalize_header(header_name)

    for cell in ws[1]:
        if normalize_header(cell.value) == target:
            return cell.column

    new_col = ws.max_column + 1
    ws.cell(row=1, column=new_col).value = header_name
    return new_col


def extract_emails(value):
    """
    从任意字符串中直接用正则提取邮箱。
    不关心它们是用 ; ； , 、 空格 或换行分隔。
    """
    if value is None:
        return []

    text = str(value)

    # Markdown / 导出文本中有时会出现 hello\@example.com
    text = text.replace(r"\@", "@")

    matches = EMAIL_RE.findall(text)

    # 保留原顺序并去重（大小写不敏感）
    result = []
    seen = set()

    for email_address in matches:
        cleaned = email_address.strip().lower()

        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            result.append(cleaned)

    return result


def parse_email_list(value):
    """
    从“已发送邮箱”或“失败邮箱”字段中恢复邮箱集合。
    """
    return set(extract_emails(value))


def append_cell_line(ws, row, col, text):
    """向某个单元格追加一行文本。"""
    old = ws.cell(row=row, column=col).value

    if old is None or not str(old).strip():
        ws.cell(row=row, column=col).value = text
    else:
        ws.cell(row=row, column=col).value = f"{old}\n{text}"


def build_html_body(raw_body):
    """
    将 Excel 中的正文转换为 Outlook HTMLBody。

    当前表格正文已有 <br>，会直接保留。
    如果以后正文是纯文本，则自动把换行变成 <br>。
    """
    if raw_body is None:
        return ""

    body = str(raw_body)

    # 清掉 Markdown 导出时可能残留的转义
    body = body.replace(r"\@", "@")
    body = body.replace(r"\*", "*")

    # 如果没有明显 HTML 标签，则按纯文本处理
    if not re.search(r"<\s*(br|p|div|ul|ol|li|strong|b|em|a)\b", body, re.I):
        body = html.escape(body)
        body = body.replace("\r\n", "\n").replace("\r", "\n")
        body = body.replace("\n", "<br>")

    return f"""\
<html>
<head>
<meta charset="utf-8">
</head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">
{body}
</body>
</html>
"""


def get_outlook_account(outlook, wanted_email):
    """
    如果指定了 SENDER_ACCOUNT，则寻找对应 Outlook 账号。
    未指定则返回 None，由 Outlook 使用默认账号。
    """
    wanted_email = (wanted_email or "").strip().lower()

    if not wanted_email:
        return None

    session = outlook.Session

    for account in session.Accounts:
        try:
            smtp = str(account.SmtpAddress).strip().lower()
        except Exception:
            smtp = ""

        if smtp == wanted_email:
            return account

    raise RuntimeError(
        f"Outlook 中没有找到发件账号：{wanted_email}\n"
        f"请检查 SENDER_ACCOUNT，或者将其留空使用 Outlook 默认账号。"
    )


def count_sent_today(ws, log_col):
    """
    从“发送记录”统计今天已经成功发了多少封。

    每一封成功邮件都会写一行：
    2026-08-26 15:30:01 | hello@example.com
    """
    today_prefix = datetime.now().strftime("%Y-%m-%d")
    count = 0

    for row in range(2, ws.max_row + 1):
        value = ws.cell(row=row, column=log_col).value

        if not value:
            continue

        for line in str(value).splitlines():
            if line.strip().startswith(today_prefix) and "|" in line:
                count += 1

    return count


def set_send_using_account(mail, account):
    """
    设置发件账号。
    pywin32 在不同 Outlook 版本上属性行为略有差异，
    因此单独封装。
    """
    if account is None:
        return

    mail.SendUsingAccount = account


def save_workbook_or_fail(wb, path):
    """
    保存 Excel。
    如果文件正被 Excel 占用，立即明确报错。
    """
    try:
        wb.save(path)
    except PermissionError:
        raise RuntimeError(
            "\n无法写入 outreach_emails.xlsx。\n"
            "请先关闭 Excel/WPS 中打开的 outreach_emails.xlsx，再重新运行 send.py。\n"
            "为了避免“邮件已经发出但状态没有保存”造成重复发送，脚本已停止。"
        )


# ============================================================
# 主程序
# ============================================================

def main():
    if sys.platform != "win32":
        raise RuntimeError("这个脚本需要在 Windows + Classic Outlook 环境运行。")

    if not XLSX_PATH.exists():
        raise FileNotFoundError(
            f"没有找到：{XLSX_PATH}\n"
            "请确认 send.py 与 outreach_emails.xlsx 位于同一个目录。"
        )

    print("=" * 72)
    print("ArkSwift Outlook Outreach Sender")
    print("=" * 72)
    print(f"Excel: {XLSX_PATH}")
    print(f"每日上限: {DAILY_LIMIT} 封")
    print(
        f"发送间隔: {MIN_DELAY_SECONDS}~{MAX_DELAY_SECONDS} 秒随机延时"
    )
    print()

    # ----------------------------------------------------------------
    # 读取 Excel
    # ----------------------------------------------------------------

    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = wb.active

    domain_col = find_column(ws, HEADER_ALIASES["domain"])
    to_col = find_column(ws, HEADER_ALIASES["to"])
    subject_col = find_column(ws, HEADER_ALIASES["subject"])
    body_col = find_column(ws, HEADER_ALIASES["body"])

    missing = []

    if domain_col is None:
        missing.append("Domain (域名)")
    if to_col is None:
        missing.append("To (收件邮箱)")
    if subject_col is None:
        missing.append("Subject (邮件主题)")
    if body_col is None:
        missing.append("Body (邮件正文)")

    if missing:
        raise RuntimeError(
            "Excel 缺少必要列："
            + ", ".join(missing)
            + "\n当前脚本不会猜测错误列，已停止。"
        )

    status_col = ensure_column(ws, "发送状态")
    sent_emails_col = ensure_column(ws, "已发送邮箱")
    failed_emails_col = ensure_column(ws, "失败邮箱")
    send_log_col = ensure_column(ws, "发送记录")
    note_col = ensure_column(ws, "发送备注")

    # 预先保存一次：
    # 1. 写入新增表头
    # 2. 确认 Excel 文件没有被锁定
    save_workbook_or_fail(wb, XLSX_PATH)

    sent_today = count_sent_today(ws, send_log_col)
    remaining_today = max(0, DAILY_LIMIT - sent_today)

    print(f"根据发送记录，今天已经发送：{sent_today} 封")
    print(f"今天还可发送：{remaining_today} 封")
    print()

    if remaining_today <= 0:
        print("今天已经达到 100 封上限，不再发送。")
        return

    # ----------------------------------------------------------------
    # 先扫描一遍，看看还有多少待发邮箱
    # ----------------------------------------------------------------

    pending_count = 0

    for row in range(2, ws.max_row + 1):
        raw_emails = ws.cell(row=row, column=to_col).value
        emails = extract_emails(raw_emails)

        sent_for_row = parse_email_list(
            ws.cell(row=row, column=sent_emails_col).value
        )

        pending_for_row = [
            e for e in emails
            if e.lower() not in sent_for_row
        ]

        pending_count += len(pending_for_row)

    print(f"表格中当前还有约 {pending_count} 个待发送邮箱。")
    print(
        f"本次最多会发送 {min(remaining_today, pending_count)} 封。"
    )
    print()

    if pending_count <= 0:
        print("没有找到需要继续发送的邮箱。")
        return

    # ----------------------------------------------------------------
    # 正式发送确认
    # ----------------------------------------------------------------

    if REQUIRE_CONFIRMATION:
        confirmation = input(
            "确认已经关闭 Excel，并准备从 Outlook 正式发送？\n"
            "请输入 SEND 后回车："
        ).strip()

        if confirmation != "SEND":
            print("未输入 SEND，本次没有发送任何邮件。")
            return

    # ----------------------------------------------------------------
    # 连接 Outlook
    # ----------------------------------------------------------------

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        send_account = get_outlook_account(outlook, SENDER_ACCOUNT)

        if send_account is None:
            print("\n发件账号：Outlook 默认账号")
        else:
            print(f"\n发件账号：{SENDER_ACCOUNT}")

        print()

        session_sent = 0

        # ------------------------------------------------------------
        # 按行处理
        # ------------------------------------------------------------

        for row in range(2, ws.max_row + 1):

            # 再次核对每日上限
            current_sent_today = count_sent_today(ws, send_log_col)

            if current_sent_today >= DAILY_LIMIT:
                print(
                    f"\n今天已达到 {DAILY_LIMIT} 封上限，自动停止。"
                )
                break

            domain = ws.cell(row=row, column=domain_col).value
            raw_emails = ws.cell(row=row, column=to_col).value
            subject = ws.cell(row=row, column=subject_col).value
            raw_body = ws.cell(row=row, column=body_col).value

            domain_text = str(domain).strip() if domain else f"第 {row} 行"

            emails = extract_emails(raw_emails)

            # --------------------------------------------------------
            # 空邮箱
            # --------------------------------------------------------

            if not emails:
                ws.cell(row=row, column=status_col).value = "无邮箱"
                ws.cell(row=row, column=note_col).value = (
                    "To 单元格为空，或未识别到有效邮箱；未发送。"
                )

                save_workbook_or_fail(wb, XLSX_PATH)

                print(f"[跳过] Row {row} | {domain_text} | 无邮箱")
                continue

            # --------------------------------------------------------
            # 缺 Subject / Body
            # --------------------------------------------------------

            if subject is None or not str(subject).strip():
                ws.cell(row=row, column=status_col).value = "缺少邮件主题"
                ws.cell(row=row, column=note_col).value = (
                    "Subject 为空；未发送。"
                )
                save_workbook_or_fail(wb, XLSX_PATH)

                print(
                    f"[跳过] Row {row} | {domain_text} | Subject 为空"
                )
                continue

            if raw_body is None or not str(raw_body).strip():
                ws.cell(row=row, column=status_col).value = "缺少邮件正文"
                ws.cell(row=row, column=note_col).value = (
                    "Body 为空；未发送。"
                )
                save_workbook_or_fail(wb, XLSX_PATH)

                print(
                    f"[跳过] Row {row} | {domain_text} | Body 为空"
                )
                continue

            # --------------------------------------------------------
            # 断点续跑：去掉这一行已经成功发过的邮箱
            # --------------------------------------------------------

            already_sent = parse_email_list(
                ws.cell(row=row, column=sent_emails_col).value
            )

            pending_emails = [
                e for e in emails
                if e.lower() not in already_sent
            ]

            if not pending_emails:
                ws.cell(row=row, column=status_col).value = "已发送"
                ws.cell(row=row, column=note_col).value = (
                    f"该行识别到 {len(emails)} 个邮箱，均已发送。"
                )

                save_workbook_or_fail(wb, XLSX_PATH)

                print(
                    f"[跳过] Row {row} | {domain_text} | 已全部发送"
                )
                continue

            # --------------------------------------------------------
            # 一个邮箱一封，逐个发送
            # --------------------------------------------------------

            html_body = build_html_body(raw_body)

            for recipient in pending_emails:

                current_sent_today = count_sent_today(ws, send_log_col)

                if current_sent_today >= DAILY_LIMIT:
                    total_sent_for_row = len(
                        parse_email_list(
                            ws.cell(
                                row=row,
                                column=sent_emails_col
                            ).value
                        )
                    )

                    if total_sent_for_row > 0:
                        ws.cell(
                            row=row,
                            column=status_col
                        ).value = "部分发送（今日上限）"
                    else:
                        ws.cell(
                            row=row,
                            column=status_col
                        ).value = "待发送（今日上限）"

                    ws.cell(row=row, column=note_col).value = (
                        f"今天已达到 {DAILY_LIMIT} 封上限；"
                        "剩余邮箱将在下次运行时继续。"
                    )

                    save_workbook_or_fail(wb, XLSX_PATH)

                    print(
                        f"\n今天达到 {DAILY_LIMIT} 封上限，自动停止。"
                    )
                    return

                try:
                    mail = outlook.CreateItem(0)
                    set_send_using_account(mail, send_account)

                    mail.To = recipient
                    mail.Subject = str(subject).strip()
                    mail.HTMLBody = html_body

                    # 正式提交给 Outlook
                    mail.Send()

                    now = datetime.now()
                    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")

                    # 成功后立即写入断点信息
                    append_cell_line(
                        ws,
                        row,
                        sent_emails_col,
                        recipient
                    )

                    append_cell_line(
                        ws,
                        row,
                        send_log_col,
                        f"{timestamp} | {recipient}"
                    )

                    # 如果这个地址以前失败过，不强行删除历史失败记录，
                    # 保留审计痕迹；发送状态以当前成功结果为准。
                    sent_after = parse_email_list(
                        ws.cell(
                            row=row,
                            column=sent_emails_col
                        ).value
                    )

                    if all(
                        e.lower() in sent_after
                        for e in emails
                    ):
                        ws.cell(
                            row=row,
                            column=status_col
                        ).value = "已发送"
                    else:
                        ws.cell(
                            row=row,
                            column=status_col
                        ).value = "部分发送"

                    ws.cell(row=row, column=note_col).value = (
                        f"识别邮箱 {len(emails)} 个；"
                        f"已发送 {len(sent_after)} 个。"
                    )

                    # 最关键的一步：
                    # 每成功一封，立刻写回 Excel。
                    save_workbook_or_fail(wb, XLSX_PATH)

                    session_sent += 1
                    sent_today_now = count_sent_today(
                        ws,
                        send_log_col
                    )

                    print(
                        f"[已发送 {sent_today_now}/{DAILY_LIMIT}] "
                        f"Row {row} | {domain_text} | {recipient}"
                    )

                except Exception as exc:
                    timestamp = datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    append_cell_line(
                        ws,
                        row,
                        failed_emails_col,
                        f"{recipient} | {timestamp} | {exc}"
                    )

                    sent_after = parse_email_list(
                        ws.cell(
                            row=row,
                            column=sent_emails_col
                        ).value
                    )

                    if sent_after:
                        ws.cell(
                            row=row,
                            column=status_col
                        ).value = "部分发送 / 有失败"
                    else:
                        ws.cell(
                            row=row,
                            column=status_col
                        ).value = "发送失败"

                    ws.cell(row=row, column=note_col).value = (
                        f"{recipient} 发送失败；下次运行会再次尝试。"
                    )

                    save_workbook_or_fail(wb, XLSX_PATH)

                    print(
                        f"[失败] Row {row} | "
                        f"{domain_text} | {recipient} | {exc}"
                    )

                    # 单个邮箱失败不让整个任务崩掉
                    # 稍作停顿后继续下一个
                    time.sleep(10)
                    continue

                # ----------------------------------------------------
                # 达到当天上限后马上停止，不再等待
                # ----------------------------------------------------

                if count_sent_today(ws, send_log_col) >= DAILY_LIMIT:
                    print(
                        f"\n今天已成功发送 {DAILY_LIMIT} 封，任务完成。"
                    )
                    return

                # ----------------------------------------------------
                # 发送节流
                # ----------------------------------------------------

                delay = random.uniform(
                    MIN_DELAY_SECONDS,
                    MAX_DELAY_SECONDS
                )

                print(
                    f"    下一封将在约 {delay:.0f} 秒后发送..."
                )
                time.sleep(delay)

                # 每 20 封多休息一次
                if (
                    session_sent > 0
                    and session_sent % BATCH_SIZE == 0
                ):
                    batch_pause = random.uniform(
                        BATCH_PAUSE_MIN_SECONDS,
                        BATCH_PAUSE_MAX_SECONDS
                    )

                    print(
                        f"\n已连续发送 {BATCH_SIZE} 封，"
                        f"额外休息约 {batch_pause:.0f} 秒...\n"
                    )

                    time.sleep(batch_pause)

        print("\n" + "=" * 72)
        print("本次表格已经处理到末尾。")
        print(
            f"今天累计成功发送："
            f"{count_sent_today(ws, send_log_col)} 封"
        )
        print("=" * 72)

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n用户手动停止。已经成功写回 Excel 的记录不会重复发送。")
    except Exception as exc:
        print("\n脚本停止：")
        print(exc)
        sys.exit(1)
