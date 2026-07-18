# backlink_opportunity_engine.py
import os
import sqlite3
import datetime
from urllib.parse import urlparse

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("请先安装 openpyxl: pip install openpyxl")

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(WORKSPACE_DIR, "backlink_intelligence.db")
OUTPUT_DIR = os.path.join(WORKSPACE_DIR, "reports")

# ==========================================
# 第一阶段：纯 Python 规则评分引擎 (Auto Scoring)
# ==========================================
def get_alpha_score(count):
    if count >= 7: return 100
    elif count >= 6: return 90
    elif count >= 5: return 80
    elif count >= 4: return 70
    elif count >= 3: return 55
    elif count >= 2: return 35
    else: return 15

def get_dr_score(dr):
    if dr >= 80: return 100
    elif dr >= 60: return 80
    elif dr >= 40: return 60
    elif dr >= 20: return 40
    else: return 20

def get_traffic_score(traffic):
    if traffic >= 100000: return 100
    elif traffic >= 50000: return 90
    elif traffic >= 10000: return 70
    elif traffic >= 1000: return 50
    else: return 20

# ==========================================
# 第二阶段：URL 意图自动识别 (Link Type)
# ==========================================
def guess_link_type(urls):
    guest_patterns = ['write-for-us', 'guest-post', 'contribute', 'submit', 'guest-blogger', 'write-for-me']
    listicle_patterns = ['best-', 'top-', '-tools', '-suppliers', 'top-10', 'top-20']
    comparison_patterns = ['-vs-', '-alternative', 'compare-', '-competitors']
    resource_patterns = ['resources', 'useful-links', 'links']
    forum_patterns = ['forum', 'thread', 'question', 'community']
    
    for url in urls:
        url_lower = str(url).lower()
        for p in guest_patterns:
            if p in url_lower: return "Guest Post (客座博客)"
        for p in comparison_patterns:
            if p in url_lower: return "Comparison (竞品对比)"
        for p in listicle_patterns:
            if p in url_lower: return "Listicle (合集/清单)"
        for p in resource_patterns:
            if p in url_lower: return "Resource Page (资源页)"
        for p in forum_patterns:
            if p in url_lower: return "Forum / Q&A (论坛问答)"
            
    return "General / Blog (常规博客/其他)"

def generate_execution_queue():
    if not os.path.exists(DB_FILE):
        print(f"[错误] 未找到数据库 {DB_FILE}，请先运行数据抓取分析脚本。")
        return

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    
    print("[进程] 正在从情报数据库中抽取数据...")
    cur.execute('''
        SELECT competitor_domain, ref_domain, ref_url, domain_rating, page_traffic, is_dofollow, is_spam, links_in_group 
        FROM backlinks 
    ''')
    rows = cur.fetchall()
    
    print("[进程] 正在通过 Python 规则引擎进行聚类和打分...")
    domain_stats = {}
    
    for row in rows:
        comp, ref_d, ref_u, dr, traf, is_dof, is_spam, links_in_group = row
        
        if ref_d not in domain_stats:
            domain_stats[ref_d] = {
                'competitors': set(),
                'urls': [],
                'max_dr': 0,
                'max_traffic': 0,
                'is_dofollow': 0,
                'is_spam': 0,
                'total_links': 0  # <--- 新增初始值
            }
            
        d = domain_stats[ref_d]
        d['competitors'].add(comp)
        d['urls'].append(ref_u)
        d['max_dr'] = max(d['max_dr'], dr)
        d['max_traffic'] = max(d['max_traffic'], traf)
        d['total_links'] += links_in_group # <--- 累加该域名下的真实外链总盘
        if is_dof == 1: d['is_dofollow'] = 1
        if is_spam == 1: d['is_spam'] = 1

    # 计算分数
    scored_domains = []
    for ref_d, d in domain_stats.items():
        alpha_count = len(d['competitors'])
        
        # 1. 基础维度算分
        score_alpha = get_alpha_score(alpha_count)
        score_dr = get_dr_score(d['max_dr'])
        score_traffic = get_traffic_score(d['max_traffic'])
        
        # 2. 权重组合 (满分100 = Alpha 35% + DR 30% + Traffic 15% + Dofollow 20%)
        base_score = (score_alpha * 0.35) + (score_dr * 0.30) + (score_traffic * 0.15)
        dofollow_bonus = 20 if d['is_dofollow'] == 1 else 0
        auto_score = round(base_score + dofollow_bonus, 1)
        
        # 3. Spam 一票否决
        if d['is_spam'] == 1:
            auto_score = -100
            
        # 4. 判断优先级
        if auto_score >= 80: priority = "P0 (核心狙击)"
        elif auto_score >= 60: priority = "P1 (重点跟进)"
        elif auto_score >= 40: priority = "P2 (日常铺垫)"
        elif auto_score > 0: priority = "P3 (长尾观察)"
        else: priority = "🗑️ SPAM (直接放弃)"
        
        # 5. 意图识别
        link_type = guess_link_type(d['urls'])
        
        # 6. 行动建议
        action_suggest = "获取邮箱并发送 Pitch"
        if "Guest Post" in link_type: action_suggest = "撰写行业文章并提交 Guest Post"
        elif "Listicle" in link_type: action_suggest = "联系作者请求加入 Top 10 List"
        elif "Comparison" in link_type: action_suggest = "请求加入产品对比评测"
        elif "Resource Page" in link_type: action_suggest = "请求将工具加入有用链接"

        scored_domains.append({
            'priority': priority,
            'domain': ref_d,
            'auto_score': auto_score,
            'link_type': link_type,
            'action': action_suggest,
            'alpha': alpha_count,
            'total_links': d['total_links'], # <--- 加上这一行
            'dr': d['max_dr'],
            'traffic': d['max_traffic'],
            'dofollow': "✅ Yes" if d['is_dofollow'] else "❌ No",
            'spam': "⚠️ Yes" if d['is_spam'] else "No",
            'example_url': d['urls'][0]
        })

    # 排序：优先度高 -> 分数高 -> DR高
    scored_domains.sort(key=lambda x: (x['auto_score'], x['dr'], x['traffic']), reverse=True)

    # ==========================================
    # 第三阶段：输出带 AI/人工 交接带的 CRM Excel
    # ==========================================
    print("[进程] 正在生成带有 AI/人工 工作流区域的最终 CRM 报表...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "外链获取执行队列 (CRM)"

    # 定义三种颜色的表头，用于物理分割不同阶段的工作流
    headers = [
        # --- Python 自动生成区 (灰底) ---
        ("A", "优先级 (Priority)", "A6A6A6"),
        ("B", "目标域名 (Domain)", "A6A6A6"),
        ("C", "🤖机器评分 (Auto Score)", "A6A6A6"),
        ("D", "内容类型 (Link Type)", "A6A6A6"),
        ("E", "建议策略 (Action Suggested)", "A6A6A6"),
        ("F", "竞品重合度 (Alpha)", "A6A6A6"),
        ("G", "🔗该域名下总外链 (Total Links)", "A6A6A6"), # <--- 新增列插在这里
        ("H", "DR权重", "A6A6A6"),
        ("I", "页面流量", "A6A6A6"),
        ("J", "Dofollow", "A6A6A6"),
        ("K", "Spam", "A6A6A6"),
        ("L", "对标链接样例", "A6A6A6"),
        
        # --- AI 预留处理区 (绿底) ---
        ("M", "🧠AI: 业务相关性评分(0-100)", "70AD47"),
        ("N", "🧠AI: 推荐创作角度(Angle)", "70AD47"),
        ("O", "🧠AI: 难度评级(Easy/Hard)", "70AD47"),
        
        # --- 人工执行区 (蓝底) ---
        ("P", "🧑‍💻人工: 最终获取难度评分(0-100)", "4472C4"),
        ("Q", "🏆总分 (Total Score) [自动计算]", "FFC000"),  
        ("R", "🧑‍💻联系邮箱 (Email)", "4472C4"),
        ("S", "🧑‍💻联系页面 (Contact URL)", "4472C4"),
        ("T", "🧑‍💻当前状态 (Status)", "4472C4")
    ]

    # 写入表头并上色
    ws.append([h[1] for h in headers])
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=h[2])
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 写入数据
    for idx, d in enumerate(scored_domains, start=2):
        row_data = [
            d['priority'],
            d['domain'],
            d['auto_score'],
            d['link_type'],
            d['action'],
            d['alpha'],
            d['total_links'],  # <--- 新字段对应 G 列
            d['dr'],
            d['traffic'],
            d['dofollow'],
            d['spam'],
            d['example_url'],
            "", # M: AI 相关性
            "", # N: AI 角度
            "", # O: AI 难度
            "", # P: 人工 难度
        ]
        ws.append(row_data)
        
        # Q列(第17列) 注入Excel公式，列号已更新：M列是AI，P列是人工
        formula = f"=(C{idx}*0.8) + (IF(ISNUMBER(M{idx}),M{idx},0)*0.15) + (IF(ISNUMBER(P{idx}),P{idx},0)*0.05)"
        ws.cell(row=idx, column=17, value=formula).font = Font(bold=True, color="C00000")
        
        # 预留最后几个空位
        ws.cell(row=idx, column=18, value="") # Email
        ws.cell(row=idx, column=19, value="") # Contact URL
        ws.cell(row=idx, column=20, value="未开始 (Not Started)") # Status

    # 调整列宽 (加入了新列 G)
    column_widths = {
        'A': 15, 'B': 25, 'C': 18, 'D': 25, 'E': 30, 
        'F': 15, 'G': 25, 'H': 10, 'I': 10, 'J': 10, 'K': 10, 'L': 40,
        'M': 25, 'N': 35, 'O': 20, 
        'P': 28, 'Q': 25, 'R': 25, 'S': 30, 'T': 20
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    out_file = os.path.join(OUTPUT_DIR, f"Backlink_Execution_Queue_{date_str}.xlsx")
    wb.save(out_file)
    
    print(f"\n[✓] 成功生成外链执行队列 (CRM级): {out_file}")
    print("--------------------------------------------------")
    print("下一步工作流建议 (SOP)：")
    print("1. [已完成] 机器已过滤掉大量低质外链，计算出了基础分 (Auto Score) 和链路类型。")
    print("2. [待执行] 将 L、M、N 列提取为 JSON 交给 GPT，批量判断业务相关性，并将数据粘回表格。")
    print("3. [待执行] 人工根据 P 列的最终得分(降序)，锁定 Top 500 开始手工寻找邮箱(Q、R列)，并推动状态流转(S列)。")
    print("--------------------------------------------------")

if __name__ == "__main__":
    generate_execution_queue()