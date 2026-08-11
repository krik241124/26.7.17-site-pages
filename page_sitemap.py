import os
import re
import glob
import logging
import requests
import pandas as pd
from urllib.parse import urlparse, unquote
import plotly.express as px
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 配置区 =================
AHREFS_DATA_DIR = "./ahrefs_data"  # 你的数据存放目录
MAX_URL_LIMIT = 2000000 
REQUEST_TIMEOUT = 15
MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE = 5

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ================= 1. 核心清洗与对齐函数 =================
def get_canonical_key(url):
    """终极 URL 归一化，确保完美对齐"""
    if not isinstance(url, str) or not url.strip(): return ""
    try:
        parsed = urlparse(url.strip().lower())
        netloc = parsed.netloc.replace('www.', '')
        path = unquote(parsed.path)
        if path.endswith('/') and len(path) > 1: path = path[:-1]
        return f"{netloc}{path}"
    except:
        return str(url).strip().lower()

# ================= 2. Sitemap 爬取 =================
class SitemapCrawler:
    def __init__(self, base_url):
        self.base_url = f"https://{base_url.replace('https://', '').replace('http://', '').strip('/')}"
        self.visited_sitemaps = set()
        self.all_urls = set()
        self.headers = {'User-Agent': 'Mozilla/5.0'}

    def get_robots_txt(self):
        robots_url = f"{self.base_url}/robots.txt"
        try:
            res = requests.get(robots_url, headers=self.headers, timeout=REQUEST_TIMEOUT, verify=False)
            if res.status_code == 200:
                sitemaps = [s.strip() for s in re.findall(r'(?i)^Sitemap:\s*(.*)', res.text, re.MULTILINE)]
                if sitemaps: return sitemaps
        except: pass
        return [f"{self.base_url}/sitemap.xml"]

    def parse_sitemap(self, sitemap_url):
        if len(self.all_urls) >= MAX_URL_LIMIT or sitemap_url in self.visited_sitemaps: return
        self.visited_sitemaps.add(sitemap_url)
        logger.info(f"正在爬取 {self.base_url} 的 Sitemap: {sitemap_url} (已收集: {len(self.all_urls)})")
        try:
            res = requests.get(sitemap_url, headers=self.headers, timeout=REQUEST_TIMEOUT, verify=False)
            if res.status_code != 200: return
            locs = re.findall(r'<loc>(.*?)</loc>', res.text)
            if '<sitemapindex' in res.text.lower() or '<sitemap>' in res.text.lower():
                for loc in locs: self.parse_sitemap(loc)
            else:
                for loc in locs:
                    if len(self.all_urls) >= MAX_URL_LIMIT: break
                    self.all_urls.add(loc)
        except: pass

    def run(self):
        logger.info(f"🚀 开始侦测域名 [{self.base_url}] 的 Sitemap 架构...")
        for sm in self.get_robots_txt(): self.parse_sitemap(sm)
        return list(self.all_urls)

# ================= 3. 动态识别域与读取 Ahrefs =================
def extract_domain_from_filename(filename):
    """尝试从 Ahrefs 导出的文件名中智能提取域名"""
    base = os.path.basename(filename)
    match = re.search(r'^([a-zA-Z0-9.\-]+?)-(?:top-pages|organic|subdomains|pages)', base, re.IGNORECASE)
    if match: return match.group(1).lower().replace('www.', '')
    return None

def scan_and_group_ahrefs_data(data_dir):
    """扫描指定目录，自动对不同域名的 Ahrefs 数据进行分组归类"""
    # 递归查找所有 CSV/XLSX
    all_files = glob.glob(os.path.join(data_dir, "**", "*.csv"), recursive=True) + \
                glob.glob(os.path.join(data_dir, "**", "*.xlsx"), recursive=True)
                
    if not all_files:
        logger.error(f"❌ 未在 {data_dir} 下找到任何 csv/xlsx 文件，请检查路径！")
        return {}

    domain_dfs = {}
    for file in all_files:
        try:
            # 兼容多种编码格式读取
            df = None
            if file.endswith('.xlsx'): df = pd.read_excel(file)
            else:
                for combo in [{'enc':'utf-16','sep':'\t'}, {'enc':'utf-8','sep':','}, {'enc':'utf-8-sig','sep':','}, {'enc':'latin1','sep':','}]:
                    try:
                        temp_df = pd.read_csv(file, encoding=combo['enc'], sep=combo['sep'], on_bad_lines='skip', low_memory=False)
                        if len(temp_df.columns) > 1: 
                            df = temp_df
                            break
                    except: continue
            
            if df is None or df.empty: continue

            # 智能提取列映射
            col_map = {}
            for col in df.columns:
                c = str(col).strip().lower()
                if c in ['url', '\ufeffurl']: col_map['URL'] = col
                elif 'traffic' in c and 'current' in c: col_map['Traffic'] = col
                elif c == 'traffic' and 'Traffic' not in col_map: col_map['Traffic'] = col
                elif 'keywords' in c and 'current' in c: col_map['Keywords'] = col
                elif c == 'keywords' and 'Keywords' not in col_map: col_map['Keywords'] = col
                elif c in ['current top keyword', 'top keyword']: col_map['Top_Keyword'] = col

            if 'URL' not in col_map: continue

            # 自动决定属于哪个域
            domain = extract_domain_from_filename(file)
            if not domain:
                first_url = df[col_map['URL']].dropna().iloc[0]
                domain = urlparse(first_url).netloc.replace('www.', '')

            if domain not in domain_dfs: domain_dfs[domain] = []
            
            clean_df = pd.DataFrame()
            clean_df['Ahrefs_URL'] = df[col_map['URL']]
            clean_df['Traffic'] = pd.to_numeric(df[col_map.get('Traffic')] if 'Traffic' in col_map else 0, errors='coerce').fillna(0)
            clean_df['Keywords'] = pd.to_numeric(df[col_map.get('Keywords')] if 'Keywords' in col_map else 0, errors='coerce').fillna(0)
            clean_df['Top_Keyword'] = df[col_map.get('Top_Keyword')] if 'Top_Keyword' in col_map else ""
            
            domain_dfs[domain].append(clean_df)
            logger.info(f"✅ 成功加载数据归入域名分组 [{domain}]: {os.path.basename(file)}")
            
        except Exception as e:
            logger.warning(f"解析 {file} 发生异常: {e}")

    return domain_dfs

# ================= 4. 构建超级目录树算法 =================
def build_tree_structure(df):
    tree = {}
    for _, row in df.iterrows():
        url = row['Final_URL']
        parsed = urlparse(url)
        domain = parsed.netloc.replace('www.', '')
        path = unquote(parsed.path).strip('/')
        
        parts = [domain] + (path.split('/') if path else [])
        
        current = tree
        for i, part in enumerate(parts):
            if part not in current:
                current[part] = {'__children__': {}, '__data__': None}
            if i == len(parts) - 1:
                current[part]['__data__'] = {
                    'url': url, 'traffic': row['Traffic'], 'kw': row['Keywords'],
                    'top_kw': row.get('Top_Keyword', ''), 'in_sitemap': row['In_Sitemap']
                }
            current = current[part]['__children__']
    return tree

def write_tree_to_txt(node_dict, f, prefix="", is_last=True, node_name=""):
    if not node_name:
        for k, v in node_dict.items():
            write_tree_to_txt(v, f, "", k == list(node_dict.keys())[-1], k)
        return

    connector = "└── " if is_last else "├── "
    data = node_dict.get('__data__')
    
    line = f"{prefix}{connector}{node_name}/"
    if data:
        if data['traffic'] > 0:
            t = int(data['traffic'])
            # 动态分级图标视觉
            if t >= 1000: icon = "🔥🔥🔥"
            elif t >= 100: icon = "🔥"
            elif t >= 10: icon = "🌟"
            else: icon = "⭐"
            
            # 防止空值显示成 nan
            kw_str = str(data['top_kw']) if pd.notna(data['top_kw']) else ""
            traffic_info = f" [{icon} 流量: {t} | 词: {int(data['kw'])} | 核心词: {kw_str}]"
        else:
            traffic_info = " [流量: 0]"
            if not data['in_sitemap']: traffic_info += " (孤岛页/不在Sitemap)"
        line = f"{prefix}{connector}{node_name}{traffic_info}"
        
    f.write(line + "\n")

    children = node_dict.get('__children__', {})
    if children:
        child_prefix = prefix + ("    " if is_last else "│   ")
        keys = list(children.keys())
        
        # 将有流量的页面顶置
        def sort_key(k):
            d = children[k].get('__data__')
            return -d['traffic'] if d else 0
            
        keys.sort(key=sort_key)
        
        traffic_keys = [k for k in keys if (children[k].get('__data__') or {}).get('traffic', 0) > 0]
        zero_keys = [k for k in keys if (children[k].get('__data__') or {}).get('traffic', 0) == 0]
        
        display_keys = traffic_keys + zero_keys[:MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE]
        hidden_count = len(zero_keys) - MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE
        
        for i, k in enumerate(display_keys):
            is_last_child = (i == len(display_keys) - 1) and (hidden_count <= 0)
            write_tree_to_txt(children[k], f, child_prefix, is_last_child, k)
            
        if hidden_count > 0:
            f.write(f"{child_prefix}└── ... (以及其他 {hidden_count} 个无流量结构页面)\n")

# ================= 5. 可视化 Treemap =================
def generate_dashboard(merged_df, domain):
    output_html = f"{domain}_seo_dashboard.html"
    def extract_path_levels(url):
        path = unquote(urlparse(url).path).strip('/')
        return path.split('/')[:3] if path else ['Home']
        
    merged_df['L1'] = merged_df['Final_URL'].apply(lambda x: extract_path_levels(x)[0] if len(extract_path_levels(x)) > 0 else 'Home')
    merged_df['L2'] = merged_df['Final_URL'].apply(lambda x: extract_path_levels(x)[1] if len(extract_path_levels(x)) > 1 else 'N/A')
    
    plot_df = merged_df[merged_df['Traffic'] > 0].copy()
    if plot_df.empty:
        logger.warning(f"⚠️ [{domain}] 没有任何流量数据，跳过树图绘制。")
        return

    plot_df['Root'] = f"{domain} 全站流量"
    fig = px.treemap(
        plot_df, path=['Root', 'L1', 'L2'], values='Traffic',
        color='Traffic', hover_data=['Keywords', 'Final_URL'],
        color_continuous_scale='Viridis', title=f"🌐 {domain} SEO 流量拓扑树图"
    )
    fig.update_layout(template="plotly_dark", margin=dict(t=50, l=25, r=25, b=25))
    fig.write_html(output_html)




def generate_html_tree(node_dict, node_name=""):
    """递归生成原生的 HTML 可折叠树菜单"""
    if not node_name:
        html = "<ul class='tree'>"
        for k, v in node_dict.items(): html += generate_html_tree(v, k)
        return html + "</ul>"

    data = node_dict.get('__data__')
    children = node_dict.get('__children__', {})
    
    label = f"📁 {node_name}" if children else f"📄 {node_name}"
    if data and data['traffic'] > 0:
        t = int(data['traffic'])
        if t >= 1000: icon, color = "🔥🔥🔥", "#ff4757"
        elif t >= 100: icon, color = "🔥", "#ffa502"
        elif t >= 10: icon, color = "🌟", "#eccc68"
        else: icon, color = "⭐", "#7bed9f"
        kw_str = str(data['top_kw']) if pd.notna(data['top_kw']) else ""
        label += f" <span style='color:{color}; font-size:14px; font-weight:bold;'>[{icon} 流量:{t} | 词:{int(data['kw'])} | 词:{kw_str}]</span>"
    elif data:
        label += f" <span style='color:#747d8c; font-size:12px;'>[流量: 0]</span>"

    html = "<li>"
    if children:
        html += f"<details open><summary>{label}</summary><ul>"
        keys = list(children.keys())
        traffic_keys = [k for k in keys if (children[k].get('__data__') or {}).get('traffic', 0) > 0]
        zero_keys = [k for k in keys if (children[k].get('__data__') or {}).get('traffic', 0) == 0]
        display_keys = traffic_keys + zero_keys[:MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE]
        
        for k in display_keys: html += generate_html_tree(children[k], k)
        if len(zero_keys) > MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE:
            html += f"<li><span style='color:#a4b0be; font-size:12px;'>... (及其他 {len(zero_keys) - MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE} 个无流量页)</span></li>"
        html += "</ul></details>"
    else:
        html += f"<div style='padding-left: 20px;'>{label}</div>"
    return html + "</li>"

# ================= 6. 主控模块 =================
def main():
    print("="*60)
    print(" 🚀 竞品 SEO 全维逆向工程分析 v4.0 (多域名全自动识别巡航版)")
    print("="*60)

    # 1. 自动读取并按域名分组
    domain_groups = scan_and_group_ahrefs_data(AHREFS_DATA_DIR)
    if not domain_groups:
        return

    # 2. 遍历每个识别到的域名，自动生成专属文件
    for domain, dfs in domain_groups.items():
        print(f"\n" + "🔥"*30)
        print(f" 开始分析目标架构: {domain.upper()}")
        print("🔥"*30 + "\n")
        
        df_ahrefs = pd.concat(dfs, ignore_index=True)
        
        # 爬取当前域的 Sitemap
        crawler = SitemapCrawler(domain)
        sitemap_urls = crawler.run()
        df_sitemap = pd.DataFrame(sitemap_urls, columns=['Sitemap_URL'])
        
        # 合并与对齐
        df_sitemap['Canonical_Key'] = df_sitemap['Sitemap_URL'].apply(get_canonical_key)
        df_ahrefs['Canonical_Key'] = df_ahrefs['Ahrefs_URL'].apply(get_canonical_key)
        
        df_sitemap = df_sitemap.drop_duplicates('Canonical_Key')
        df_ahrefs = df_ahrefs.sort_values('Traffic', ascending=False).drop_duplicates('Canonical_Key')
        
        merged_df = pd.merge(df_sitemap, df_ahrefs, on='Canonical_Key', how='outer')
        merged_df['Final_URL'] = merged_df['Ahrefs_URL'].combine_first(merged_df['Sitemap_URL'])
        merged_df['Traffic'] = merged_df.get('Traffic', 0).fillna(0)
        merged_df['Keywords'] = merged_df.get('Keywords', 0).fillna(0)
        merged_df['In_Sitemap'] = merged_df['Sitemap_URL'].notna()

        # 生成独家 TXT 树状图
        logger.info(f"正在渲染 {domain} 全站拓扑目录树 (因为有数百万页面折叠，这可能需要几十秒)...")
        tree_data = build_tree_structure(merged_df)
        tree_file = f"{domain}_seo_architecture_tree.txt"
        with open(tree_file, "w", encoding="utf-8") as f:
            f.write(f"🌐 {domain} SEO 拓扑全景图\n")
            f.write("="*60 + "\n")
            write_tree_to_txt(tree_data, f)
        
        # 生成大屏与数据表格
        generate_dashboard(merged_df, domain)
        


        # ====== [新增] 追加酷炫的可折叠 HTML 树状图到底部 ======
        html_tree = generate_html_tree(tree_data)
        with open(f"{domain}_seo_dashboard.html", "a", encoding="utf-8") as f:
            f.write(f"""
            <hr style='border:1px solid #444; margin: 40px 0;'>
            <h2 style='color:#fff; text-align:center; font-family:sans-serif;'>🌳 {domain} 详细目录拓扑树 (原生可折叠)</h2>
            <div style='background:#1e1e1e; color:#d4d4d4; padding:20px; font-family:monospace; line-height:1.8; overflow-x: auto;'>
                <style>
                    .tree details summary {{ cursor: pointer; outline: none; list-style: none; }}
                    .tree details summary::-webkit-details-marker {{ display: none; }}
                    .tree details summary:before {{ content: '▶ '; color: #00a8ff; font-size: 12px; }}
                    .tree details[open] summary:before {{ content: '▼ '; }}
                    .tree ul {{ list-style-type: none; padding-left: 25px; border-left: 1px dashed #555; margin-top: 5px; }}
                </style>
                {html_tree}
            </div>
            """)




        excel_file = f"{domain}_seo_metrics.xlsx"
        merged_df['Dir_L1'] = merged_df['Final_URL'].apply(lambda x: urlparse(x).path.strip('/').split('/')[0] if urlparse(x).path.strip('/') else 'Home')
        summary = merged_df.groupby('Dir_L1').agg(Pages=('Final_URL','count'), Traffic=('Traffic','sum')).sort_values('Traffic', ascending=False)
        
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name='目录流量汇总')
            top_traffic = merged_df[merged_df['Traffic'] > 0].sort_values('Traffic', ascending=False)
            if top_traffic.empty: top_traffic = merged_df.head(1000) # 防止无流量时出错
            top_traffic.to_excel(writer, sheet_name='有流量页面清单', index=False)
            
        print(f"\n🎉 [{domain}] 分析完成！已生成专属文件: ")
        print(f" 🌳 拓扑图: {tree_file}")
        print(f" 📊 大屏图: {domain}_seo_dashboard.html")
        print(f" 📈 报表库: {excel_file}\n")

if __name__ == "__main__":
    main()