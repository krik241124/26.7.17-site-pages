import pandas as pd
import os
import re

# ================= 配置区域 =================
INPUT_FILE = 'input/targets.xlsx'
OUTPUT_FILE = 'output/outreach_emails.xlsx'
TEMPLATE_FILE = 'templates/template_01.txt'
# ============================================

def ensure_directories():
    """确保必要的文件夹存在"""
    for folder in ['input', 'output', 'templates']:
        if not os.path.exists(folder):
            os.makedirs(folder)

def load_template(filepath):
    """读取邮件模板，分离 Subject 和 Body"""
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 找到第一行非空的文本作为 Subject
    subject = ""
    body_lines = []
    
    for i, line in enumerate(lines):
        if line.strip() and not subject:
            subject = line.strip()
            # 剩下的部分作为正文 (去掉紧接着的空行)
            body_lines = lines[i+1:]
            break
            
    # 去除正文开头可能多余的空行
    body = "".join(body_lines).lstrip()
    return subject, body

def get_company_name_from_domain(domain):
    """从域名中提取公司名 (例如: uscreen.tv -> Uscreen)"""
    if pd.isna(domain) or not isinstance(domain, str):
        return "Partner"
    
    # 去除 www. 
    clean_domain = domain.replace('www.', '').strip()
    # 根据 '.' 分割，取第一部分并首字母大写
    company_name = clean_domain.split('.')[0].capitalize()
    return company_name

def main():
    ensure_directories()
    
    # 1. 检查文件是否存在
    if not os.path.exists(TEMPLATE_FILE):
        print(f"❌ 找不到模板文件: {TEMPLATE_FILE}")
        return
    if not os.path.exists(INPUT_FILE):
        print(f"❌ 找不到输入文件: {INPUT_FILE}")
        return

    print("✅ 正在读取模板...")
    subject_template, body_template = load_template(TEMPLATE_FILE)

    print("✅ 正在读取目标数据...")
    try:
        # 1. 先尝试读取
        df = pd.read_excel(INPUT_FILE)
        
        # 2. 检查当前有没有表头 (通过判断列名里有没有包含 Domain 或 域名)
        has_header = any('domain' in str(c).lower() or '域名' in str(c) for c in df.columns)
        
        if not has_header:
            print("⚠️ 检测到 targets.xlsx 缺失表头，正在自动插入并保存...")
            
            # 无表头重新读取，确保第一行数据(比如 uscreen.tv) 不被当成表头吃掉
            df = pd.read_excel(INPUT_FILE, header=None)
            
            # 准备你定义的 29 个标准表头
            standard_headers = [
                "目标域名 (Domain)", "分层与优先级 (Tier)", "🤖机器评分 (Auto Score)", "内容类型 (Link Type)",
                "建议策略 (Action Suggested)", "竞品重合度 (Alpha)", "🔗该域名下总外链 (Total Links)", "DR权重",
                "页面流量", "Dofollow", "Spam", "对标链接样例", "🧠AI：业务相关性评分 (0-100)",
                "🧠AI：配套内容 (如 Pitch 邮件|大纲)", "🧠AI：难度评级 (Easy|Hard)", "🧑‍💻联系邮箱 (Email)",
                "🧑‍💻获取入口URL", "🏆总分 (Total Score)（自动计算）", "🧑‍💻人工：审阅状态", "🧑‍💻当前状态 (Status)",
                "Last Action Date", "Link Result", "🤝合作模式 (Cooperation Model)", "💰合作报价 (Cost/Price)",
                "✍️内容提供方 (Content Provider)", "📍发布板块与位置 (Placement)", "⏳链接留存期 (Link Duration)",
                "🔗约定链接属性 (Agreed Attribute)", "身份"
            ]
            
            # 容错：确保定义的表头数量跟实际数据的列数对得上
            if len(df.columns) > len(standard_headers):
                standard_headers.extend([f"未知扩展列_{i}" for i in range(len(standard_headers), len(df.columns))])
            else:
                standard_headers = standard_headers[:len(df.columns)]
                
            # 赋予表头
            df.columns = standard_headers
            
            # 直接重写原文件，把加上表头的数据存回去
            df.to_excel(INPUT_FILE, index=False)
            print("✅ 表头已成功插入，并直接更新保存到了 targets.xlsx 中！")
            
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    # 寻找包含 "域名" 或 "Domain" 的列，以及包含 "邮箱" 或 "Email" 的列
    domain_col = next((col for col in df.columns if 'domain' in str(col).lower() or '域名' in str(col)), None)
    email_col = next((col for col in df.columns if 'email' in str(col).lower() or '邮箱' in str(col)), None)

    if not domain_col:
        print("❌ 在 Excel 中找不到表示域名的列！请检查表头。")
        return
    if not email_col:
        print("❌ 在 Excel 中找不到表示邮箱的列！请检查表头。")
        return

    output_data = []

    print("✅ 正在生成邮件内容...")
    # 遍历每一行数据
    for index, row in df.iterrows():
        domain = row[domain_col]
        email = row[email_col]
        
        # 处理空域名的情况
        if pd.isna(domain):
            continue
            
        # 处理空邮箱的情况
        if pd.isna(email) or str(email).strip() == "":
            email = "" # 留空方便后续手动补全

        # 提取名字并替换模板
        company_name = get_company_name_from_domain(domain)
        customized_body = body_template.replace("[Name]", company_name)
        
        # 组装一条干净的数据
        output_data.append({
            "Domain (域名)": domain,
            "To (收件邮箱)": email,
            "Subject (邮件主题)": subject_template,
            "Body (邮件正文)": customized_body
        })

    # 2. 转换为 DataFrame
    out_df = pd.DataFrame(output_data)

    print("✅ 正在导出排版友好的 Excel 文件...")
    # 3. 使用 xlsxwriter 引擎导出，以便设置列宽和自动换行 (这对你手动复制非常重要)
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
    out_df.to_excel(writer, index=False, sheet_name='Outreach Emails')
    
    # 获取 xlsxwriter 的 workbook 和 worksheet 对象
    workbook = writer.book
    worksheet = writer.sheets['Outreach Emails']
    
    # 设置格式：垂直居中、正文自动换行
    wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter'})
    normal_format = workbook.add_format({'valign': 'vcenter'})
    
    # 设置列宽
    worksheet.set_column('A:A', 20, normal_format) # Domain
    worksheet.set_column('B:B', 30, normal_format) # To (Email)
    worksheet.set_column('C:C', 60, normal_format) # Subject
    worksheet.set_column('D:D', 100, wrap_format)  # Body (设置较宽且自动换行)

    writer.close()
    
    print(f"🎉 成功！共生成了 {len(out_df)} 封邮件，已保存至: {OUTPUT_FILE}")
    print("👉 提示：打开 output 文件夹下的 Excel 文件，双击 Body 列即可直接 Ctrl+A, Ctrl+C 复制到 Outlook！")

if __name__ == "__main__":
    main()
