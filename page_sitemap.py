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
    total_rows = len(df)
    
    # 🌟 修改点：用 enumerate 获取真实的循环次数 count
    for count, (_, row) in enumerate(df.iterrows()):
        
        # ========== 实时进度反馈 ==========
        if count > 0 and count % 50000 == 0:
            logger.info(f"   👉 拓扑树构建进度: 已处理 {count} / {total_rows} 页面 ({(count/total_rows)*100:.1f}%)")
        # ===============================
        
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

    # === 【之前新增的：递归计算包含子节点的总流量和总页面数】 ===
    def calculate_node_stats(node):
        data = node.get('__data__')
        node_traffic = data.get('traffic', 0) if data else 0
        node_pages = 1 if data else 0
        node_kw = data.get('kw', 0) if data else 0
        
        # 核心词逻辑
        best_kw_str = data.get('top_kw', '') if data else ''
        if pd.isna(best_kw_str): best_kw_str = ''
        max_traffic_for_kw = node_traffic if best_kw_str else -1
        
        for k, child in node.get('__children__', {}).items():
            c_traffic, c_pages, c_kw, c_top_kw, c_max_t = calculate_node_stats(child)
            node_traffic += c_traffic
            node_pages += c_pages
            node_kw += c_kw
            if c_max_t > max_traffic_for_kw and c_top_kw:
                best_kw_str = c_top_kw
                max_traffic_for_kw = c_max_t
                
        node['__total_traffic__'] = node_traffic
        node['__total_pages__'] = node_pages
        node['__total_kw__'] = node_kw
        node['__best_kw__'] = best_kw_str
        node['__max_traffic_for_kw__'] = max_traffic_for_kw
        
        return node_traffic, node_pages, node_kw, best_kw_str, max_traffic_for_kw

    for k, v in tree.items():
        calculate_node_stats(v)
    return tree

def write_tree_to_txt(node_dict, f, prefix="", is_last=True, node_name=""):
    if not node_name:
        for k, v in node_dict.items():
            write_tree_to_txt(v, f, "", k == list(node_dict.keys())[-1], k)
        return

    connector = "└── " if is_last else "├── "
    data = node_dict.get('__data__')
    total_traffic = node_dict.get('__total_traffic__', 0)
    total_pages = node_dict.get('__total_pages__', 0)
    
    line = f"{prefix}{connector}{node_name} ({total_pages}个页面)"
    
    if total_traffic > 0:
        t = int(total_traffic)
        if t >= 1000: icon = "🔥🔥🔥"
        elif t >= 100: icon = "🔥"
        elif t >= 10: icon = "🌟"
        else: icon = "⭐"
        
        # 兼容处理有无具体节点数据的情况
        kw_val = int(data['kw']) if data else 0
        kw_str = str(data['top_kw']) if data and pd.notna(data.get('top_kw')) else ""
        traffic_info = f" [{icon} 流量: {t} | 词: {kw_val} | 核心词: {kw_str}]"
    else:
        traffic_info = " [流量: 0]"
        if data and not data.get('in_sitemap'): traffic_info += " (孤岛页/不在Sitemap)"
        
    f.write(line + traffic_info + "\n")

    children = node_dict.get('__children__', {})
    if children:
        child_prefix = prefix + ("    " if is_last else "│   ")
        keys = list(children.keys())
        
        # 核心修复：按包含子目录在内的总流量进行降序排列
        keys.sort(key=lambda k: -children[k].get('__total_traffic__', 0))
        
        traffic_keys = [k for k in keys if children[k].get('__total_traffic__', 0) > 0]
        zero_keys = [k for k in keys if children[k].get('__total_traffic__', 0) == 0]
        
        display_keys = traffic_keys + zero_keys[:MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE]
        hidden_count = len(zero_keys) - MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE
        
        for i, k in enumerate(display_keys):
            is_last_child = (i == len(display_keys) - 1) and (hidden_count <= 0)
            write_tree_to_txt(children[k], f, child_prefix, is_last_child, k)
            
        if hidden_count > 0:
            f.write(f"{child_prefix}└── ... (以及其他 {hidden_count} 个无流量结构页面)\n")
            

# ================= 5. 可视化 Treemap =================
def generate_dashboard(merged_df, domain, tree_data):
    output_html = f"{domain}_seo_dashboard.html"
    def extract_path_levels(url):
        path = unquote(urlparse(url).path).strip('/')
        return path.split('/')[:3] if path else ['Home']
        
    merged_df['L1'] = merged_df['Final_URL'].apply(lambda x: extract_path_levels(x)[0] if len(extract_path_levels(x)) > 0 else '根目录')
    merged_df['L2'] = merged_df['Final_URL'].apply(lambda x: extract_path_levels(x)[1] if len(extract_path_levels(x)) > 1 else '[无二级目录]')
    merged_df['L3'] = merged_df['Final_URL'].apply(lambda x: extract_path_levels(x)[2] if len(extract_path_levels(x)) > 2 else '[无三级目录]')
    
    plot_df = merged_df[merged_df['Traffic'] > 0].copy()
    if plot_df.empty: return

    plot_df['Root'] = f"{domain} 全站流量"
    plot_df['Top_Keyword'] = plot_df['Top_Keyword'].fillna('无')
    
    fig = px.treemap(
        plot_df, path=['Root', 'L1', 'L2', 'L3'], values='Traffic',
        color='Traffic', color_continuous_scale='Viridis',
        title=f"🌐 {domain} SEO 流量拓扑树图 (Sitemap 总收录: {len(merged_df)} 页面)"
    )
    
    # 强制注入准确统计数据以对齐下方目录树
    import numpy as np
    if fig.data:
        ids = fig.data[0].ids
        new_customdata = []
        for node_id in ids:
            parts = str(node_id).split('/')
            if len(parts) > 0 and parts[0] == f"{domain} 全站流量":
                parts = parts[1:]
            
            # 定位到对应目录节点
            root_key = list(tree_data.keys())[0] if tree_data else None
            current = tree_data.get(root_key, {}) if root_key else {}
            
            for p in parts:
                if p in ('根目录', '[无二级目录]', '[无三级目录]', 'Home'): continue
                if current and '__children__' in current and p in current['__children__']:
                    current = current['__children__'][p]
                else:
                    current = {}
                    break
            
            # 把精确获取到的词数和汇总数据怼进图表里
            if current:
                t_pages = current.get('__total_pages__', 0)
                t_kw = current.get('__total_kw__', 0)
                b_kw = current.get('__best_kw__', '无')
                if not b_kw: b_kw = '无'
                new_customdata.append([node_id, t_kw, b_kw, t_pages])
            else:
                new_customdata.append([node_id, 0, '无', 0])
                
        fig.data[0].customdata = np.array(new_customdata)
        
    fig.update_traces(
        hovertemplate=(
            "<b style='font-size:14px;'>节点: %{label}</b><br>"
            "<hr>"
            "📈 <b>流量总计 (含子页)</b>: %{value:.0f}<br>"
            "🎯 <b>关键词数 (含子页)</b>: %{customdata[1]}<br>"
            "👑 <b>核心搜索词 (含子页)</b>: %{customdata[2]}<br>"
            "📄 <b>包含页面数 (含子页)</b>: %{customdata[3]}"
        )
    )
    
    fig.update_layout(template="plotly_dark", margin=dict(t=50, l=25, r=25, b=25))
    fig.write_html(output_html)




def generate_html_tree(node_dict, node_name=None, progress=None, level=0):
    if progress is None:
        progress = {"count": 0}

    progress["count"] += 1
    if progress["count"] % 20000 == 0:
        logger.info(f"   👉 网页渲染进度: 绝对没卡死！已生成 {progress['count']} 个树形节点的 HTML 代码...")

    if node_name is None:
        html = "<ul class='tree' id='seo-tree'>"
        for k, v in node_dict.items(): html += generate_html_tree(v, k, progress, level=1)
        return html + "</ul>"

    data = node_dict.get('__data__')
    children = node_dict.get('__children__', {})
    
    total_traffic = node_dict.get('__total_traffic__', 0)
    total_pages = node_dict.get('__total_pages__', 0)
    
    label = f"📁 {node_name}" if children else f"📄 {node_name}"
    label += f" <span style='color:#888; font-size:12px;'>({total_pages}个页面)</span>"

    if total_traffic > 0:
        t = int(total_traffic)
        if t >= 1000: icon, color = "🔥🔥🔥", "#ff4757"
        elif t >= 100: icon, color = "🔥", "#ffa502"
        elif t >= 10: icon, color = "🌟", "#eccc68"
        else: icon, color = "⭐", "#7bed9f"
        
        # 兼容读取父级合并上来的核心关键词
        kw_str = str(data['top_kw']) if data and pd.notna(data.get('top_kw')) else ""
        if not kw_str and node_dict.get('__best_kw__'): kw_str = node_dict.get('__best_kw__')
        kw_show = f" | 核心词:{kw_str}" if kw_str else ""
        
        label += f" <span style='color:{color}; font-size:14px; font-weight:bold;'>[{icon} 流量:{t}{kw_show}]</span>"
    else:
        label += f" <span style='color:#747d8c; font-size:12px;'>[流量: 0]</span>"

    html = "<li>"
    if children:
        # ✅ 去掉了 open 属性，改为注入等级属性供 JS 控制（彻底防崩溃）
        html += f"<details data-level='{level}' data-traffic='{total_traffic}'><summary>{label}</summary><ul>"
        keys = list(children.keys())
        
        traffic_keys = [k for k in keys if children[k].get('__total_traffic__', 0) > 0]
        zero_keys = [k for k in keys if children[k].get('__total_traffic__', 0) == 0]
        traffic_keys.sort(key=lambda k: -children[k].get('__total_traffic__', 0))
        
        display_keys = traffic_keys + zero_keys[:MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE]
        hidden_keys = zero_keys[MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE:]
        
        for k in display_keys: html += generate_html_tree(children[k], k, progress, level + 1)
        
        if hidden_keys:
            html += f"<li><details data-level='{level+1}' data-traffic='0' name='zombie_accordion'><summary style='color:#a4b0be; font-size:12px; cursor:pointer;'>... (点击展开其他 {len(hidden_keys)} 个无流量页面)</summary><ul>"
            for k in hidden_keys: html += generate_html_tree(children[k], k, progress, level + 2)
            html += "</ul></details></li>"
            
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
        logger.info(f"⏳ 步骤 1/4: 正在解析 {len(merged_df)} 个URL并构建节点字典 (大约需要 20-40 秒)...")
        tree_data = build_tree_structure(merged_df)
        
        logger.info(f"⏳ 步骤 2/4: 正在生成纯文本拓扑树 TXT 文件...")
        tree_file = f"{domain}_seo_architecture_tree.txt"
        with open(tree_file, "w", encoding="utf-8") as f:
            f.write(f"🌐 {domain} SEO 拓扑全景图\n")
            f.write("="*60 + "\n")
            write_tree_to_txt(tree_data, f)
        
        # 注意这里：把 tree_data 给传进图表里强行获取真实目录数据
        logger.info(f"⏳ 步骤 3/4: 正在生成可视化数据大屏 HTML...")
        generate_dashboard(merged_df, domain, tree_data) 
        
        logger.info(f"⏳ 步骤 4/4: 正在递归生成可折叠的底层 HTML 树形代码...")

        # ====== 追加酷炫的可折叠 HTML 树状图到底部（含JS按钮面板防卡） ======
        html_tree = generate_html_tree(tree_data)
        with open(f"{domain}_seo_dashboard.html", "a", encoding="utf-8") as f:
            f.write(f"""
            <hr style='border:1px solid #444; margin: 40px 0;'>
            <h2 style='color:#fff; text-align:center; font-family:sans-serif;'>🌳 {domain} 详细目录拓扑树 (原生可折叠)</h2>
            
            <!-- JS 交互面板 -->
            <div style="text-align:center; margin-bottom: 20px;">
                <button onclick="expandTree(2)" style="padding:10px 15px; margin:5px; cursor:pointer; background:#00a8ff; color:#fff; border:none; border-radius:4px; font-weight:bold;">展开到 1 级目录</button>
                <button onclick="expandTree(3)" style="padding:10px 15px; margin:5px; cursor:pointer; background:#00a8ff; color:#fff; border:none; border-radius:4px; font-weight:bold;">展开到 2 级目录</button>
                <button onclick="expandTree(4)" style="padding:10px 15px; margin:5px; cursor:pointer; background:#00a8ff; color:#fff; border:none; border-radius:4px; font-weight:bold;">展开到 3 级目录</button>
                <button onclick="collapseAll()" style="padding:10px 15px; margin:5px; cursor:pointer; background:#444; color:#fff; border:none; border-radius:4px; font-weight:bold;">全部折叠</button>
            </div>
            
            <script>
                function expandTree(targetLevel) {{
                    document.body.style.cursor = 'wait';
                    setTimeout(() => {{
                        const details = document.querySelectorAll('#seo-tree details');
                        details.forEach(d => {{
                            const level = parseInt(d.getAttribute('data-level')) || 999;
                            const traffic = parseFloat(d.getAttribute('data-traffic')) || 0;
                            
                            if (level < targetLevel) {{
                                d.open = true;
                            }} else {{
                                d.open = false;
                            }}
                            
                            // 🚀 防崩溃降级策略：如果你请求打开过深的层级，只要这分支没流量，强制继续让他折叠！
                            if (traffic === 0 && level >= 2) {{
                                d.open = false;
                            }}
                        }});
                        document.body.style.cursor = 'default';
                    }}, 10);
                }}
                
                function collapseAll() {{
                    const details = document.querySelectorAll('#seo-tree details');
                    details.forEach(d => d.open = false);
                }}
                
                // 页面打开时，仅默认展示出 1级目录，防止 200万个节点全放出来让电脑死机
                window.onload = function() {{ expandTree(2); }};
            </script>

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
