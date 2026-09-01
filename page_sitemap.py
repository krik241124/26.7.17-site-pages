# page-sitemap.py
# Ahrefs / Semrush + Sitemap SEO Architecture Analyzer
# Ahrefs Top Pages supports current exports with AI responses columns; extra columns are ignored safely.
# 页面数据源自动识别：Ahrefs 优先，Semrush 作为兼容回退。
# Sitemap 对齐、目录树、TXT、Treemap、Excel 与分片懒加载共用同一套后续逻辑。

import os
import re
import csv
import glob
import json
import math
import shutil
import hashlib
import logging
import itertools
from urllib.parse import urlparse, unquote

import requests
import pandas as pd
import plotly.graph_objects as go
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ================= 配置区 =================
WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(WORKSPACE_DIR, "data")

# Sitemap 最多收集 URL 数。保留旧版 200 万上限。
MAX_URL_LIMIT = 2_000_000
REQUEST_TIMEOUT = 15

# TXT 中每个目录最多展示多少个“零流量结构页”，防止 TXT 无限膨胀。
MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE = 5

# HTML 分片懒加载：每次只把一个目录的 N 个直属子节点加载进 DOM。
# 300 对百万级站点比较稳妥；需要更少文件可调大，需要更低浏览器内存可调小。
LAZY_CHILD_PAGE_SIZE = 300

# 深层目录折叠后自动释放已加载 DOM，再次展开时从磁盘分片重新加载。
UNLOAD_DEEP_TREE_ON_COLLAPSE = True
UNLOAD_FROM_LEVEL = 2

SEM_RUSH_MARKER = "-organic.pagesv3-"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


# ================= 1. 基础清洗 / 数据源对齐 =================
def clean_header(value):
    return str(value or "").replace("\ufeff", "").strip()


def normalize_header(value):
    return re.sub(r"[\s_]+", " ", clean_header(value).lower()).strip()


def detect_encoding(file_path):
    candidates = ["utf-8-sig", "utf-8", "utf-16", "cp1252", "latin1"]
    for enc in candidates:
        try:
            with open(file_path, "r", encoding=enc) as f:
                f.read(4096)
            return enc
        except UnicodeError:
            continue
    return "utf-8-sig"


def discover_csv_dialect(file_path, encoding):
    if file_path.lower().endswith(".tsv"):
        return "\t"

    try:
        with open(file_path, "r", encoding=encoding, newline="") as f:
            sample = f.read(8192)
        return csv.Sniffer().sniff(sample, delimiters=",;\t").delimiter
    except Exception:
        return ","


def extract_competitor_from_filename(filename):
    """
    例：
    sindebella.com-organic.PagesV3-us-20260821-2026-08-22T14_50_41Z.csv
    -> sindebella.com
    """
    base = os.path.basename(filename)
    lower = base.lower()

    marker_pos = lower.find(SEM_RUSH_MARKER)
    if marker_pos != -1:
        return base[:marker_pos].strip().lower().replace("www.", "")

    match = re.match(r"^(.*?)-organic\.pagesv3(?:-|_)", base, flags=re.I)
    if match:
        return match.group(1).strip().lower().replace("www.", "")

    return None


def extract_semrush_database_from_filename(filename):
    """从 Semrush Pages V3 文件名中提取 us / uk / de / fr 等数据库。"""
    base = os.path.basename(filename)
    match = re.search(r"-organic\.pagesv3-([a-z]{2,3})-", base, flags=re.I)
    if match:
        return match.group(1).lower()
    return "unknown"


def get_canonical_key(url):
    """沿用旧版 URL 归一化：域名去 www、path 解码、忽略 query/fragment、去尾斜杠。"""
    if not isinstance(url, str) or not url.strip():
        return ""
    try:
        parsed = urlparse(url.strip().lower())
        netloc = parsed.netloc.replace("www.", "")
        path = unquote(parsed.path)
        if path.endswith("/") and len(path) > 1:
            path = path[:-1]
        return f"{netloc}{path}"
    except Exception:
        return str(url).strip().lower()


def numeric_series(series, default=0.0):
    if series is None:
        return default
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False),
        errors="coerce"
    ).fillna(default)


# ================= 2. Sitemap 爬取 =================
class SitemapCrawler:
    def __init__(self, base_url):
        cleaned = base_url.replace("https://", "").replace("http://", "").strip("/")
        self.base_url = f"https://{cleaned}"
        self.visited_sitemaps = set()
        self.all_urls = set()
        self.headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0 Safari/537.36"
            )
        }
        self.session = requests.Session()
        self.session.headers.update(self.headers)

    def get_robots_txt(self):
        robots_url = f"{self.base_url}/robots.txt"
        try:
            res = self.session.get(
                robots_url,
                timeout=REQUEST_TIMEOUT,
                verify=False
            )
            if res.status_code == 200:
                sitemaps = [
                    s.strip()
                    for s in re.findall(
                        r"(?i)^Sitemap:\s*(.*)",
                        res.text,
                        re.MULTILINE
                    )
                ]
                if sitemaps:
                    return sitemaps
        except Exception as e:
            logger.debug(f"robots.txt 读取失败: {e}")

        return [f"{self.base_url}/sitemap.xml"]

    def parse_sitemap(self, sitemap_url):
        if len(self.all_urls) >= MAX_URL_LIMIT:
            return
        if sitemap_url in self.visited_sitemaps:
            return

        self.visited_sitemaps.add(sitemap_url)
        logger.info(
            f"正在爬取 {self.base_url} 的 Sitemap: {sitemap_url} "
            f"(已收集: {len(self.all_urls):,})"
        )

        try:
            res = self.session.get(
                sitemap_url,
                timeout=REQUEST_TIMEOUT,
                verify=False
            )
            if res.status_code != 200:
                return

            text = res.text
            locs = re.findall(r"<loc[^>]*>(.*?)</loc>", text, flags=re.I | re.S)
            is_index = "<sitemapindex" in text.lower()

            if is_index:
                for loc in locs:
                    loc = loc.strip()
                    if loc:
                        self.parse_sitemap(loc)
            else:
                for loc in locs:
                    if len(self.all_urls) >= MAX_URL_LIMIT:
                        break
                    loc = loc.strip()
                    if loc:
                        self.all_urls.add(loc)

        except Exception as e:
            logger.debug(f"Sitemap 读取失败 [{sitemap_url}]: {e}")

    def run(self):
        logger.info(f"🚀 开始侦测域名 [{self.base_url}] 的 Sitemap 架构...")
        for sm in self.get_robots_txt():
            self.parse_sitemap(sm)

        if len(self.all_urls) >= MAX_URL_LIMIT:
            logger.warning(
                f"⚠️ 已达到 MAX_URL_LIMIT={MAX_URL_LIMIT:,}，后续 Sitemap URL 已停止收集。"
            )

        # 直接返回 set，避免这里再复制出一份同体量 list。
        return self.all_urls


# ================= 3. 动态识别并读取 Ahrefs / Semrush 页面数据 =================
def extract_ahrefs_domain_from_filename(filename):
    """
    兼容 Ahrefs 常见导出文件名，例如：
    www.vevor.com-top-pages-subdomains-all--co_2026-08-14_14-58-27.csv
    -> vevor.com
    """
    base = os.path.basename(filename)
    match = re.match(
        r"^([a-zA-Z0-9.\-]+?)-(?:top-pages(?:-subdomains)?|organic|subdomains|pages)",
        base,
        flags=re.I
    )
    if match:
        return match.group(1).lower().replace("www.", "")
    return None


def infer_domain_from_urls(df, url_col):
    non_null_urls = df[url_col].dropna()
    for value in non_null_urls:
        value = str(value).strip()
        if not value or value.lower() == "nan":
            continue
        domain = urlparse(value).netloc.lower().replace("www.", "")
        if domain:
            return domain
    return None


def read_tabular_export(file_path):
    """
    读取 CSV / TSV / XLSX。

    空导出的定义包括：
    - 0 字节文件；
    - 只有 BOM / 空白字符；
    - pandas 抛 EmptyDataError；
    - 只有空列/Unnamed 列且没有任何有效单元格。

    这些情况统一返回 0x0 DataFrame，交给扫描层依据文件名识别数据源。
    """
    lower = file_path.lower()

    # 最稳定的第一层：真正的 0 字节文件根本不交给 pandas。
    try:
        if os.path.getsize(file_path) == 0:
            return pd.DataFrame()
    except OSError:
        pass

    # 文本文件如果只有 BOM / CRLF / 空格，也按空导出处理。
    if lower.endswith((".csv", ".tsv")):
        try:
            with open(file_path, "rb") as raw_f:
                raw_sample = raw_f.read(8192)
            text_sample = raw_sample.decode("utf-8-sig", errors="ignore")
            if not text_sample.strip():
                return pd.DataFrame()
        except OSError:
            pass

    try:
        if lower.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_path)
        else:
            encoding = detect_encoding(file_path)
            delimiter = discover_csv_dialect(file_path, encoding)
            df = pd.read_csv(
                file_path,
                encoding=encoding,
                sep=delimiter,
                on_bad_lines="skip",
                low_memory=False
            )
    except (pd.errors.EmptyDataError, pd.errors.ParserError):
        return pd.DataFrame()

    # 某些“空 CSV”实际是 ,,,,, 或 Excel 导出的全空白网格。
    # pandas 会造出 Unnamed 列，此处进一步归零。
    if df is None:
        return pd.DataFrame()

    meaningful_headers = [
        clean_header(c) for c in df.columns
        if clean_header(c) and not clean_header(c).lower().startswith("unnamed:")
    ]

    if not meaningful_headers:
        if df.empty:
            return pd.DataFrame()
        try:
            has_value = df.apply(
                lambda col: col.astype(str).str.strip().replace("nan", "").ne("").any()
            ).any()
        except Exception:
            has_value = False
        if not has_value:
            return pd.DataFrame()

    return df




def empty_normalized_page_df():
    """统一的 0 行页面数据结构；后续 Sitemap 合并逻辑可直接复用。"""
    return pd.DataFrame({
        "Source_URL": pd.Series(dtype="object"),
        "Traffic": pd.Series(dtype="float64"),
        "Keywords": pd.Series(dtype="float64"),
        "Top_Keyword": pd.Series(dtype="object"),
    })


def infer_empty_export_from_filename(file_path):
    """
    仅用于真正的 0 行/0 字节导出：没有表头可识别时，以文件名兜底。

    注意：空文件没有表头，因此先识别 Semrush 明确的
    ``-organic.PagesV3-`` 标记；否则再按 Ahrefs 文件名规则判断。
    这不改变正常文件的 Ahrefs -> Semrush 表头识别优先级。
    """
    semrush_domain = extract_competitor_from_filename(file_path)
    if semrush_domain:
        return {
            "domain": semrush_domain,
            "source": "semrush",
            "database": extract_semrush_database_from_filename(file_path),
            "df": empty_normalized_page_df(),
        }

    ahrefs_domain = extract_ahrefs_domain_from_filename(file_path)
    if ahrefs_domain:
        return {
            "domain": ahrefs_domain,
            "source": "ahrefs",
            "database": None,
            "df": empty_normalized_page_df(),
        }

    return None

def find_column(normalized_to_original, *candidates):
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized_to_original:
            return normalized_to_original[key]
    return None


def detect_page_export_source(normalized_headers):
    """
    数据源识别顺序固定为 Ahrefs -> Semrush。
    返回 "ahrefs" / "semrush" / None。
    """
    # Ahrefs Top Pages / Top pages by organic traffic 常见结构。
    ahrefs_has_url = "url" in normalized_headers
    ahrefs_has_traffic = any(
        x in normalized_headers
        for x in {"current traffic", "traffic current"}
    )
    ahrefs_has_keywords = any(
        x in normalized_headers
        for x in {
            "current # of keywords",
            "current number of keywords",
            "current keywords",
            "keywords current",
        }
    )

    if ahrefs_has_url and ahrefs_has_traffic and ahrefs_has_keywords:
        return "ahrefs"

    # Semrush Organic Research > Pages V3。
    semrush_required = {"url", "traffic", "number of keywords"}
    semrush_signature = {
        "traffic (%)",
        "top keyword",
        "primary intent",
        "positions with informational intents in top 20",
    }
    if (
        semrush_required.issubset(normalized_headers)
        and bool(normalized_headers & semrush_signature)
    ):
        return "semrush"

    return None


def normalize_ahrefs_export(df, file_path, normalized_to_original):
    url_col = find_column(normalized_to_original, "URL")
    traffic_col = find_column(
        normalized_to_original,
        "Current traffic",
        "Traffic current",
        "Traffic"
    )
    kw_col = find_column(
        normalized_to_original,
        "Current # of keywords",
        "Current number of keywords",
        "Current keywords",
        "Keywords current",
        "Keywords"
    )
    top_kw_col = find_column(
        normalized_to_original,
        "Current top keyword",
        "Top keyword"
    )

    if url_col is None or traffic_col is None or kw_col is None:
        return None

    domain = extract_ahrefs_domain_from_filename(file_path)
    if not domain:
        domain = infer_domain_from_urls(df, url_col)
    if not domain:
        return None

    clean_df = pd.DataFrame()
    clean_df["Source_URL"] = df[url_col].astype(str).str.strip()
    clean_df["Traffic"] = numeric_series(df[traffic_col], 0.0)
    clean_df["Keywords"] = numeric_series(df[kw_col], 0.0)
    clean_df["Top_Keyword"] = (
        df[top_kw_col].fillna("").astype(str)
        if top_kw_col is not None
        else ""
    )

    clean_df = clean_df[
        clean_df["Source_URL"].notna()
        & (clean_df["Source_URL"].str.len() > 0)
        & (clean_df["Source_URL"].str.lower() != "nan")
    ]

    return {
        "domain": domain,
        "source": "ahrefs",
        "database": None,
        "df": clean_df,
    }


def normalize_semrush_export(df, file_path, normalized_to_original):
    url_col = find_column(normalized_to_original, "URL")
    traffic_col = find_column(normalized_to_original, "Traffic")
    kw_col = find_column(normalized_to_original, "Number of Keywords")
    top_kw_col = find_column(
        normalized_to_original,
        "Top Keyword",
        "Main Keyword"
    )

    if url_col is None or traffic_col is None or kw_col is None:
        return None

    domain = extract_competitor_from_filename(file_path)
    if not domain:
        domain = infer_domain_from_urls(df, url_col)
    if not domain:
        return None

    semrush_db = extract_semrush_database_from_filename(file_path)

    clean_df = pd.DataFrame()
    clean_df["Source_URL"] = df[url_col].astype(str).str.strip()
    clean_df["Traffic"] = numeric_series(df[traffic_col], 0.0)
    clean_df["Keywords"] = numeric_series(df[kw_col], 0.0)
    clean_df["Top_Keyword"] = (
        df[top_kw_col].fillna("").astype(str)
        if top_kw_col is not None
        else ""
    )

    clean_df = clean_df[
        clean_df["Source_URL"].notna()
        & (clean_df["Source_URL"].str.len() > 0)
        & (clean_df["Source_URL"].str.lower() != "nan")
    ]

    return {
        "domain": domain,
        "source": "semrush",
        "database": semrush_db,
        "df": clean_df,
    }


def scan_and_group_page_data(data_dir):
    """
    扫描 data/ 中的页面导出文件并统一成：
      Source_URL / Traffic / Keywords / Top_Keyword

    优先级：
      1. 单文件先按 Ahrefs 表头识别；不是 Ahrefs 才尝试 Semrush。
      2. 同一域名若同时存在 Ahrefs 与 Semrush，整个域名使用 Ahrefs，忽略 Semrush。
      3. 不同域名可以在同一次运行中分别使用 Ahrefs 或 Semrush。
    """
    patterns = [
        "*.csv", "*.tsv", "*.CSV", "*.TSV",
        "*.xlsx", "*.xls", "*.XLSX", "*.XLS"
    ]
    all_files = []
    for pattern in patterns:
        all_files.extend(
            glob.glob(os.path.join(data_dir, "**", pattern), recursive=True)
        )
    all_files = sorted(set(all_files))

    if not all_files:
        logger.error(f"未在 {data_dir} 下找到 csv/tsv/xlsx/xls 文件。")
        return {}

    ahrefs_by_domain = {}
    semrush_by_domain_db = {}

    for file_path in all_files:
        # 先拿文件名候选。这样即使底层解析器对完全空文件报错，
        # 仍然能把已知 Ahrefs / Semrush 命名的文件视为 0 流量数据源。
        filename_fallback = infer_empty_export_from_filename(file_path)

        try:
            try:
                df = read_tabular_export(file_path)
            except Exception as read_error:
                if filename_fallback is not None:
                    logger.info(
                        f"页面数据文件无可解析内容，按 0 流量处理: "
                        f"{os.path.basename(file_path)} ({type(read_error).__name__})"
                    )
                    df = pd.DataFrame()
                else:
                    raise

            if df is None:
                continue

            # 0 字节 / 无表头 / 全空白表：无法靠列判断，按文件名兜底识别。
            if len(df.columns) == 0:
                result = filename_fallback
                if result is None:
                    logger.info(
                        f"跳过空且无法识别来源的文件: {os.path.basename(file_path)}"
                    )
                    continue

                domain = result["domain"]
                if result["source"] == "ahrefs":
                    ahrefs_by_domain.setdefault(domain, []).append(result["df"])
                    logger.info(
                        f"已识别 Ahrefs [{domain}]: {os.path.basename(file_path)} "
                        f"(0 行，按 0 流量处理)"
                    )
                else:
                    semrush_db = result["database"]
                    semrush_by_domain_db.setdefault(
                        (domain, semrush_db), []
                    ).append(result["df"])
                    logger.info(
                        f"已识别 Semrush [{domain}] [DB={semrush_db}]: "
                        f"{os.path.basename(file_path)} (0 行，按 0 流量处理)"
                    )
                continue

            normalized_to_original = {
                normalize_header(c): c for c in df.columns
            }
            normalized_headers = set(normalized_to_original.keys())
            source = detect_page_export_source(normalized_headers)

            if source == "ahrefs":
                result = normalize_ahrefs_export(
                    df,
                    file_path,
                    normalized_to_original
                )
                if not result:
                    continue
                domain = result["domain"]
                ahrefs_by_domain.setdefault(domain, []).append(result["df"])
                logger.info(
                    f"已识别 Ahrefs [{domain}]: {os.path.basename(file_path)} "
                    f"({len(result['df']):,} 行" +
                    ("，按 0 流量处理)" if result["df"].empty else ")")
                )
                continue

            if source == "semrush":
                result = normalize_semrush_export(
                    df,
                    file_path,
                    normalized_to_original
                )
                if not result:
                    continue
                domain = result["domain"]
                semrush_db = result["database"]
                semrush_by_domain_db.setdefault(
                    (domain, semrush_db), []
                ).append(result["df"])
                logger.info(
                    f"已识别 Semrush [{domain}] [DB={semrush_db}]: "
                    f"{os.path.basename(file_path)} ({len(result['df']):,} 行" +
                    ("，按 0 流量处理)" if result["df"].empty else ")")
                )
                continue

            logger.info(
                f"跳过未识别的页面数据文件: {os.path.basename(file_path)}"
            )

        except Exception as e:
            logger.warning(f"解析 {file_path} 发生异常: {e}")

    grouped = {}

    # Ahrefs 先入组；这是域名级优先级。
    for domain, dfs in ahrefs_by_domain.items():
        grouped[(domain, "ahrefs", None)] = dfs

    # 只有该域名没有 Ahrefs 时，才使用 Semrush。
    for (domain, semrush_db), dfs in semrush_by_domain_db.items():
        if domain in ahrefs_by_domain:
            logger.info(
                f"[{domain}] 同时存在 Ahrefs 与 Semrush 页面数据；按优先级使用 Ahrefs，"
                f"忽略 Semrush DB={semrush_db}。"
            )
            continue
        grouped[(domain, "semrush", semrush_db)] = dfs

    if not grouped:
        logger.error(
            "已找到数据文件，但未识别出 Ahrefs Top Pages 或 Semrush Organic Pages V3。"
        )

    return grouped


# ================= 4. 构建超级目录树算法 =================
def build_tree_structure(df):
    tree = {}
    total_rows = len(df)

    for count, row in enumerate(df.itertuples(index=False), start=1):
        if count % 50_000 == 0:
            logger.info(
                f"   👉 拓扑树构建进度: 已处理 {count:,} / {total_rows:,} 页面 "
                f"({(count / max(total_rows, 1)) * 100:.1f}%)"
            )

        url = getattr(row, "Final_URL")
        if not isinstance(url, str) or not url:
            continue

        parsed = urlparse(url)
        domain = parsed.netloc.replace("www.", "")
        path = unquote(parsed.path).strip("/")
        parts = [domain] + (path.split("/") if path else [])

        current = tree
        for i, part in enumerate(parts):
            if part not in current:
                current[part] = {
                    "__children__": {},
                    "__data__": None
                }

            if i == len(parts) - 1:
                top_kw = getattr(row, "Top_Keyword", "")
                if pd.isna(top_kw):
                    top_kw = ""

                current[part]["__data__"] = {
                    "url": url,
                    "traffic": float(getattr(row, "Traffic", 0) or 0),
                    "kw": float(getattr(row, "Keywords", 0) or 0),
                    "top_kw": str(top_kw),
                    "in_sitemap": bool(getattr(row, "In_Sitemap", False))
                }

            current = current[part]["__children__"]

    def calculate_node_stats(node):
        data = node.get("__data__")
        node_traffic = data.get("traffic", 0) if data else 0
        node_pages = 1 if data else 0
        node_kw = data.get("kw", 0) if data else 0

        best_kw_str = data.get("top_kw", "") if data else ""
        if pd.isna(best_kw_str):
            best_kw_str = ""
        max_traffic_for_kw = node_traffic if best_kw_str else -1

        for child in node.get("__children__", {}).values():
            c_traffic, c_pages, c_kw, c_top_kw, c_max_t = calculate_node_stats(child)
            node_traffic += c_traffic
            node_pages += c_pages
            node_kw += c_kw
            if c_max_t > max_traffic_for_kw and c_top_kw:
                best_kw_str = c_top_kw
                max_traffic_for_kw = c_max_t

        node["__total_traffic__"] = node_traffic
        node["__total_pages__"] = node_pages
        node["__total_kw__"] = node_kw
        node["__best_kw__"] = best_kw_str
        node["__max_traffic_for_kw__"] = max_traffic_for_kw

        return node_traffic, node_pages, node_kw, best_kw_str, max_traffic_for_kw

    for root_node in tree.values():
        calculate_node_stats(root_node)

    return tree


def write_tree_to_txt(node_dict, f, prefix="", is_last=True, node_name=""):
    """保留旧版 TXT 行为：零流量结构页每个目录只展示前 N 个。"""
    if not node_name:
        root_keys = list(node_dict.keys())
        for i, key in enumerate(root_keys):
            write_tree_to_txt(
                node_dict[key],
                f,
                "",
                i == len(root_keys) - 1,
                key
            )
        return

    connector = "└── " if is_last else "├── "
    data = node_dict.get("__data__")
    total_traffic = node_dict.get("__total_traffic__", 0)
    total_pages = node_dict.get("__total_pages__", 0)

    line = f"{prefix}{connector}{node_name} ({total_pages:,}个页面)"

    if total_traffic > 0:
        t = int(total_traffic)
        if t >= 1000:
            icon = "🔥🔥🔥"
        elif t >= 100:
            icon = "🔥"
        elif t >= 10:
            icon = "🌟"
        else:
            icon = "⭐"

        kw_val = int(data["kw"]) if data else 0
        kw_str = str(data["top_kw"]) if data and data.get("top_kw") else ""
        traffic_info = f" [{icon} 流量: {t:,} | 词: {kw_val:,} | 核心词: {kw_str}]"
    else:
        traffic_info = " [流量: 0]"
        if data and not data.get("in_sitemap"):
            traffic_info += " (孤岛页/不在Sitemap)"

    f.write(line + traffic_info + "\n")

    children = node_dict.get("__children__", {})
    if children:
        child_prefix = prefix + ("    " if is_last else "│   ")
        keys = list(children.keys())
        keys.sort(key=lambda k: -children[k].get("__total_traffic__", 0))

        traffic_keys = [
            k for k in keys
            if children[k].get("__total_traffic__", 0) > 0
        ]
        zero_keys = [
            k for k in keys
            if children[k].get("__total_traffic__", 0) == 0
        ]

        display_keys = traffic_keys + zero_keys[:MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE]
        hidden_count = max(
            0,
            len(zero_keys) - MAX_ZOMBIE_PAGES_PER_DIR_IN_TREE
        )

        for i, key in enumerate(display_keys):
            is_last_child = (
                i == len(display_keys) - 1
                and hidden_count <= 0
            )
            write_tree_to_txt(
                children[key],
                f,
                child_prefix,
                is_last_child,
                key
            )

        if hidden_count > 0:
            f.write(
                f"{child_prefix}└── ... "
                f"(以及其他 {hidden_count:,} 个无流量结构页面)\n"
            )


# ================= 5. Treemap =================
def generate_dashboard(domain_label, output_html, tree_data):
    """保留旧版逻辑：Treemap 只绘制有流量分支，避免 Sitemap 百万零流量页撑爆图表。"""
    plot_data = {
        "ids": [],
        "labels": [],
        "parents": [],
        "values": [],
        "kw": [],
        "best_kw": [],
        "pages": [],
        "urls": []
    }

    def flatten_tree_for_plotly(node, parent_id, node_id, label):
        traffic = node.get("__total_traffic__", 0)
        if traffic <= 0:
            return

        pages = node.get("__total_pages__", 0)
        kw = node.get("__total_kw__", 0)
        best_kw = node.get("__best_kw__", "无") or "无"

        data = node.get("__data__")
        if data and data.get("url"):
            real_url = data.get("url")
        else:
            real_url = f"[目录组] https://{node_id}"

        plot_data["ids"].append(node_id)
        plot_data["labels"].append(label)
        plot_data["parents"].append(parent_id)
        plot_data["values"].append(traffic)
        plot_data["kw"].append(kw)
        plot_data["best_kw"].append(best_kw)
        plot_data["pages"].append(pages)
        plot_data["urls"].append(real_url)

        for child_key, child_node in node.get("__children__", {}).items():
            child_id = f"{node_id}/{child_key}" if node_id else child_key
            flatten_tree_for_plotly(child_node, node_id, child_id, child_key)

    for root_key, root_node in tree_data.items():
        flatten_tree_for_plotly(root_node, "", root_key, root_key)

    if not plot_data["ids"]:
        logger.info(f"{domain_label} 页面数据流量为 0；跳过流量 Treemap，继续生成 Sitemap 目录树。")
        with open(output_html, "w", encoding="utf-8") as f:
            f.write(
                "<!doctype html><html><head><meta charset='utf-8'>"
                "<title>SEO Architecture</title></head>"
                "<body style='margin:0;background:#111;color:#fff;font-family:sans-serif;'>"
                f"<h1 style='text-align:center;padding:30px;'>{domain_label} SEO 流量架构</h1>"
                "<p style='text-align:center;color:#aaa;'>页面数据中没有有流量节点，下面仍可查看 Sitemap 目录树。</p>"
                "</body></html>"
            )
        return

    fig = go.Figure(go.Treemap(
        ids=plot_data["ids"],
        labels=plot_data["labels"],
        parents=plot_data["parents"],
        values=plot_data["values"],
        branchvalues="total",
        customdata=list(zip(
            plot_data["kw"],
            plot_data["best_kw"],
            plot_data["pages"],
            plot_data["urls"]
        )),
        hovertemplate=(
            "<b style='font-size:16px; color:#00a8ff;'>节点: %{label}</b><br>"
            "<hr style='border-color:#444;'>"
            "🔗 <b>完整链接</b>: %{customdata[3]}<br>"
            "📈 <b>流量总计 (含子页)</b>: %{value:.0f}<br>"
            "🎯 <b>关键词数 (含子页)</b>: %{customdata[0]}<br>"
            "👑 <b>核心搜索词 (含子页)</b>: %{customdata[1]}<br>"
            "📄 <b>包含页面数 (含子页)</b>: %{customdata[2]}<br>"
            "<extra></extra>"
        ),
        textinfo="label+value",
        marker=dict(
            colors=plot_data["values"],
            colorscale="Viridis",
            showscale=True
        )
    ))

    fig.update_layout(
        title=f"🌐 {domain_label} SEO 流量架构 1:1 透视拓扑图",
        template="plotly_dark",
        margin=dict(t=60, l=10, r=10, b=10)
    )
    fig.write_html(output_html)


# ================= 6. HTML 分片懒加载目录树 =================
def stable_chunk_id(parent_chunk_id, child_name):
    raw = f"{parent_chunk_id}\x1f{child_name}".encode("utf-8", errors="ignore")
    return "n_" + hashlib.sha1(raw).hexdigest()[:20]


def node_to_lazy_record(name, node, child_chunk_id=None):
    data = node.get("__data__") or {}
    return {
        "name": name,
        "has_children": bool(node.get("__children__")),
        "chunk_id": child_chunk_id,
        "traffic": node.get("__total_traffic__", 0),
        "pages": node.get("__total_pages__", 0),
        "kw": node.get("__total_kw__", 0),
        "best_kw": node.get("__best_kw__", "") or "",
        "url": data.get("url", "") or "",
        "in_sitemap": bool(data.get("in_sitemap", False)) if data else True,
        "has_data": bool(data)
    }


def write_lazy_tree_assets(tree_data, asset_dir):
    """
    将完整树拆成很多小 JS 数据块。

    关键点：
    - 主 HTML 不包含完整树；
    - 一个目录一次只加载 LAZY_CHILD_PAGE_SIZE 个直属子节点；
    - 使用动态 <script src=...> 而不是 fetch，因此直接双击 file:// HTML 也能工作；
    - 未展开目录对应的 JS 分片不会进入浏览器内存 / DOM。
    """
    if os.path.exists(asset_dir):
        shutil.rmtree(asset_dir)
    os.makedirs(asset_dir, exist_ok=True)

    root_chunk_id = "root_" + hashlib.sha1(
        os.path.abspath(asset_dir).encode("utf-8")
    ).hexdigest()[:16]

    # virtual root: 与普通节点统一处理
    stack = [(root_chunk_id, tree_data, 0)]
    total_nodes_written = 0
    total_files_written = 0

    while stack:
        parent_chunk_id, children, level = stack.pop()
        if not children:
            continue

        # 延续旧版优先逻辑：有流量分支优先，零流量分支随后。
        traffic_keys = [
            k for k, v in children.items()
            if v.get("__total_traffic__", 0) > 0
        ]
        zero_keys = [
            k for k, v in children.items()
            if v.get("__total_traffic__", 0) <= 0
        ]
        traffic_keys.sort(
            key=lambda k: -children[k].get("__total_traffic__", 0)
        )

        ordered_iter = itertools.chain(traffic_keys, zero_keys)
        page_index = 0
        buffer = []

        def flush_buffer(records, page_no, is_last_page):
            nonlocal total_files_written, total_nodes_written
            payload = {
                "chunk_id": parent_chunk_id,
                "page": page_no,
                "has_more": not is_last_page,
                "nodes": records
            }
            output_path = os.path.join(
                asset_dir,
                f"{parent_chunk_id}_{page_no:05d}.js"
            )
            with open(output_path, "w", encoding="utf-8") as out:
                out.write(
                    "window.__seoReceiveChunk("
                    + json.dumps(parent_chunk_id, ensure_ascii=False)
                    + ","
                    + str(page_no)
                    + ","
                    + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
                    + ");"
                )
            total_files_written += 1
            total_nodes_written += len(records)

        # 为了知道最后一页，需要看直属子节点总数。
        total_children = len(children)
        processed = 0

        for key in ordered_iter:
            child = children[key]
            child_children = child.get("__children__", {})
            child_chunk_id = None

            if child_children:
                child_chunk_id = stable_chunk_id(parent_chunk_id, key)
                stack.append((child_chunk_id, child_children, level + 1))

            buffer.append(
                node_to_lazy_record(
                    key,
                    child,
                    child_chunk_id=child_chunk_id
                )
            )
            processed += 1

            if len(buffer) >= LAZY_CHILD_PAGE_SIZE:
                is_last = processed >= total_children
                flush_buffer(buffer, page_index, is_last)
                buffer = []
                page_index += 1

        if buffer:
            flush_buffer(buffer, page_index, True)

        if total_nodes_written and total_nodes_written % 50_000 < LAZY_CHILD_PAGE_SIZE:
            logger.info(
                f"   👉 懒加载分片生成进度: 已写入 {total_nodes_written:,} 个节点 / "
                f"{total_files_written:,} 个分片文件"
            )

    # 空站点 / Sitemap 未发现 URL 时也写一个空 root chunk。
    # 这样 Dashboard 打开根节点时不会请求一个不存在的 JS 文件。
    if total_files_written == 0:
        payload = {
            "chunk_id": root_chunk_id,
            "page": 0,
            "has_more": False,
            "nodes": []
        }
        output_path = os.path.join(
            asset_dir,
            f"{root_chunk_id}_00000.js"
        )
        with open(output_path, "w", encoding="utf-8") as out:
            out.write(
                "window.__seoReceiveChunk("
                + json.dumps(root_chunk_id, ensure_ascii=False)
                + ",0,"
                + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
                + ");"
            )
        total_files_written = 1

    manifest = {
        "root_chunk_id": root_chunk_id,
        "page_size": LAZY_CHILD_PAGE_SIZE,
        "total_nodes": total_nodes_written,
        "total_chunk_files": total_files_written,
        "unload_on_collapse": UNLOAD_DEEP_TREE_ON_COLLAPSE,
        "unload_from_level": UNLOAD_FROM_LEVEL
    }
    with open(
        os.path.join(asset_dir, "manifest.json"),
        "w",
        encoding="utf-8"
    ) as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    logger.info(
        f"✅ HTML 树已拆分为懒加载资源: {total_nodes_written:,} 个节点 / "
        f"{total_files_written:,} 个 JS 分片"
    )
    return root_chunk_id, manifest


def append_lazy_tree_to_dashboard(
    output_html,
    domain_label,
    asset_dir_name,
    root_chunk_id,
    manifest,
    root_total_pages=0,
    root_total_traffic=0,
    data_source_label=""
):
    unload_js = "true" if UNLOAD_DEEP_TREE_ON_COLLAPSE else "false"

    # 注意：这里故意不嵌入 tree_data；HTML 只有一个空壳，节点按点击动态加载。
    html = f"""
<hr style='border:1px solid #444; margin:40px 0;'>
<section id='lazy-tree-section' style='background:#111;color:#d4d4d4;padding:20px;font-family:monospace;line-height:1.8;'>
    <h2 style='color:#fff;text-align:center;font-family:sans-serif;'>{domain_label} 目录拓扑树</h2>
    <p style='text-align:center;color:#888;font-family:sans-serif;'>
        Load mode: chunked on demand; batch size: {LAZY_CHILD_PAGE_SIZE}; source: {data_source_label}.
    </p>

    <div style='text-align:center;margin-bottom:20px;font-family:sans-serif;'>
        <button onclick='expandTree(1)' class='tree-btn'>展开到 1 级目录</button>
        <button onclick='expandTree(2)' class='tree-btn'>展开到 2 级目录</button>
        <button onclick='expandTree(3)' class='tree-btn'>展开到 3 级目录</button>
        <button onclick='collapseAll()' class='tree-btn tree-btn-dark'>全部折叠</button>
    </div>

    <div id='tree-status' style='text-align:center;color:#777;font-family:sans-serif;margin-bottom:12px;'>
        Tree nodes: {manifest['total_nodes']:,}; chunk files: {manifest['total_chunk_files']:,}.
    </div>

    <ul class='tree' id='seo-tree'>
        <li>
            <details id='tree-root' open data-level='0' data-traffic='{root_total_traffic}' data-loaded='0' data-chunk-id='{root_chunk_id}'>
                <summary><span class='folder-label'>🌐 {domain_label}</span> <span class='node-meta'>({root_total_pages:,}个页面)</span></summary>
                <ul class='children'></ul>
            </details>
        </li>
    </ul>
</section>

<style>
    #lazy-tree-section .tree-btn {{
        padding:10px 15px;margin:5px;cursor:pointer;background:#00a8ff;color:#fff;
        border:none;border-radius:4px;font-weight:bold;
    }}
    #lazy-tree-section .tree-btn-dark {{ background:#444; }}
    #lazy-tree-section .tree {{ list-style:none;padding-left:0; }}
    #lazy-tree-section .tree ul {{
        list-style:none;padding-left:25px;border-left:1px dashed #555;margin-top:5px;
    }}
    #lazy-tree-section details > summary {{ cursor:pointer;outline:none;list-style:none; }}
    #lazy-tree-section details > summary::-webkit-details-marker {{ display:none; }}
    #lazy-tree-section details > summary:before {{ content:'▶ ';color:#00a8ff;font-size:12px; }}
    #lazy-tree-section details[open] > summary:before {{ content:'▼ '; }}
    #lazy-tree-section .node-meta {{ color:#888;font-size:12px; }}
    #lazy-tree-section .node-hot {{ font-size:14px;font-weight:bold; }}
    #lazy-tree-section .leaf {{ padding-left:20px; }}
    #lazy-tree-section .load-more {{
        margin:6px 0 10px 4px;padding:6px 10px;cursor:pointer;background:#2f3542;
        color:#dfe4ea;border:1px solid #57606f;border-radius:4px;
    }}
    #lazy-tree-section .loading {{ color:#70a1ff;font-size:12px;padding-left:20px; }}
    #lazy-tree-section .load-error {{ color:#ff6b81;font-size:12px;padding-left:20px; }}
</style>

<script>
(function() {{
    const ASSET_BASE = './{asset_dir_name}/';
    const ROOT_CHUNK_ID = {json.dumps(root_chunk_id)};
    const UNLOAD_ON_COLLAPSE = {unload_js};
    const UNLOAD_FROM_LEVEL = {UNLOAD_FROM_LEVEL};
    const pending = new Map();
    const loadedScripts = new Set();

    function esc(value) {{
        return String(value ?? '')
            .replaceAll('&', '&amp;')
            .replaceAll('<', '&lt;')
            .replaceAll('>', '&gt;')
            .replaceAll('"', '&quot;')
            .replaceAll("'", '&#039;');
    }}

    window.__seoReceiveChunk = function(chunkId, page, payload) {{
        const key = chunkId + ':' + page;
        const waiter = pending.get(key);
        if (waiter) {{
            waiter.resolve(payload);
            pending.delete(key);
        }}
    }};

    function loadChunk(chunkId, page) {{
        const key = chunkId + ':' + page;
        if (pending.has(key)) return pending.get(key).promise;

        let resolveFn, rejectFn;
        const promise = new Promise((resolve, reject) => {{
            resolveFn = resolve;
            rejectFn = reject;
        }});
        pending.set(key, {{promise, resolve: resolveFn, reject: rejectFn}});

        const script = document.createElement('script');
        const src = ASSET_BASE + chunkId + '_' + String(page).padStart(5, '0') + '.js';
        script.src = src;
        script.async = true;
        script.onload = () => {{
            loadedScripts.add(src);
            // 执行完即可移除 script 标签，数据已交给 JS，不让 script DOM 越积越多。
            script.remove();
        }};
        script.onerror = () => {{
            const waiter = pending.get(key);
            if (waiter) {{
                waiter.reject(new Error('无法加载懒加载分片: ' + src));
                pending.delete(key);
            }}
            script.remove();
        }};
        document.head.appendChild(script);
        return promise;
    }}

    function badgeHtml(n) {{
        const pages = Number(n.pages || 0);
        const traffic = Number(n.traffic || 0);
        const bestKw = n.best_kw ? ' | 核心词:' + esc(n.best_kw) : '';
        let html = " <span class='node-meta'>(" + pages.toLocaleString() + "个页面)</span>";

        if (traffic > 0) {{
            let icon = '⭐', color = '#7bed9f';
            if (traffic >= 1000) {{ icon = '🔥🔥🔥'; color = '#ff4757'; }}
            else if (traffic >= 100) {{ icon = '🔥'; color = '#ffa502'; }}
            else if (traffic >= 10) {{ icon = '🌟'; color = '#eccc68'; }}

            html += " <span class='node-hot' style='color:" + color + ";'>[" + icon +
                    " 流量:" + Math.round(traffic).toLocaleString() + bestKw + "]</span>";
        }} else {{
            html += " <span class='node-meta'>[流量: 0]</span>";
            if (n.has_data && !n.in_sitemap) {{
                html += " <span class='node-meta'>(孤岛页/不在Sitemap)</span>";
            }}
        }}
        return html;
    }}

    function makeNode(n, level) {{
        const li = document.createElement('li');
        const icon = n.has_children ? '📁 ' : '📄 ';
        const label = icon + esc(n.name) + badgeHtml(n);

        if (n.has_children) {{
            const details = document.createElement('details');
            details.dataset.level = String(level);
            details.dataset.traffic = String(n.traffic || 0);
            details.dataset.loaded = '0';
            details.dataset.chunkId = n.chunk_id;

            const summary = document.createElement('summary');
            summary.innerHTML = label;
            const ul = document.createElement('ul');
            ul.className = 'children';

            details.appendChild(summary);
            details.appendChild(ul);
            details.addEventListener('toggle', () => onToggle(details));
            li.appendChild(details);
        }} else {{
            const div = document.createElement('div');
            div.className = 'leaf';
            div.innerHTML = label;
            li.appendChild(div);
        }}
        return li;
    }}

    function removeLoadMore(ul) {{
        const old = ul.querySelector(':scope > .load-more');
        if (old) old.remove();
    }}

    async function loadPage(details, page) {{
        const ul = details.querySelector(':scope > ul.children');
        const chunkId = details.dataset.chunkId;
        removeLoadMore(ul);

        const loading = document.createElement('li');
        loading.className = 'loading';
        loading.textContent = '正在按需加载...';
        ul.appendChild(loading);

        try {{
            const payload = await loadChunk(chunkId, page);
            loading.remove();

            const level = Number(details.dataset.level || 0) + 1;
            const frag = document.createDocumentFragment();
            for (const node of payload.nodes || []) {{
                frag.appendChild(makeNode(node, level));
            }}
            ul.appendChild(frag);

            details.dataset.loaded = '1';
            details.dataset.nextPage = String(page + 1);

            if (payload.has_more) {{
                const btn = document.createElement('button');
                btn.className = 'load-more';
                btn.textContent = '加载下一批 {LAZY_CHILD_PAGE_SIZE} 个直属节点';
                btn.onclick = async (e) => {{
                    e.preventDefault();
                    e.stopPropagation();
                    btn.disabled = true;
                    await loadPage(details, Number(details.dataset.nextPage || 0));
                }};
                ul.appendChild(btn);
            }}
        }} catch (err) {{
            loading.className = 'load-error';
            loading.textContent = '加载失败：' + err.message + '。请确认 HTML 与资源目录保持在同一位置。';
        }}
    }}

    async function ensureLoaded(details) {{
        if (details.dataset.loaded === '1') return;
        await loadPage(details, 0);
    }}

    async function onToggle(details) {{
        if (details.open) {{
            await ensureLoaded(details);
            return;
        }}

        const level = Number(details.dataset.level || 0);
        if (UNLOAD_ON_COLLAPSE && level >= UNLOAD_FROM_LEVEL) {{
            // 折叠后稍后释放子 DOM；如果用户马上又展开则不清理。
            setTimeout(() => {{
                if (!details.open) {{
                    const ul = details.querySelector(':scope > ul.children');
                    if (ul) ul.replaceChildren();
                    details.dataset.loaded = '0';
                    details.dataset.nextPage = '0';
                }}
            }}, 250);
        }}
    }}

    window.expandTree = async function(targetLevel) {{
        document.body.style.cursor = 'wait';
        const root = document.getElementById('tree-root');
        const queue = [root];

        while (queue.length) {{
            const d = queue.shift();
            if (!d) continue;
            const level = Number(d.dataset.level || 0);
            const traffic = Number(d.dataset.traffic || 0);

            if (level < targetLevel) {{
                // 延续旧版防崩逻辑：深层零流量分支不做批量强制展开。
                if (!(traffic === 0 && level >= 2)) {{
                    d.open = true;
                    await ensureLoaded(d);
                }}
            }} else {{
                d.open = false;
            }}

            if (d.open) {{
                const direct = d.querySelectorAll(':scope > ul.children > li > details');
                for (const child of direct) queue.push(child);
            }}
        }}
        document.body.style.cursor = 'default';
    }};

    window.collapseAll = function() {{
        document.querySelectorAll('#seo-tree details').forEach(d => {{
            if (d.id !== 'tree-root') d.open = false;
        }});
        const root = document.getElementById('tree-root');
        if (root) root.open = false;
    }};

    // 初始只加载根目录第一批直属节点；不加载任何未展开深层目录。
    // 脚本本身位于树容器之后，因此无需等待 window.onload，避免追加到 Plotly HTML 后错过 load 事件。
    setTimeout(async () => {{
        const root = document.getElementById('tree-root');
        if (root) await ensureLoaded(root);
    }}, 0);
}})();
</script>
"""

    with open(output_html, "a", encoding="utf-8") as f:
        f.write(html)


# ================= 7. 主控模块 =================
def main():
    print("=" * 68)
    print(" Page Sitemap SEO Architecture Analyzer v6.3")
    print(" Ahrefs priority / Semrush fallback / lazy-loaded sitemap tree")
    print("=" * 68)

    domain_groups = scan_and_group_page_data(DATA_DIR)
    if not domain_groups:
        return

    # 同一 domain 复用 Sitemap 抓取结果，避免不同 Semrush DB 重复联网。
    sitemap_cache = {}

    for (domain, data_source, source_db), dfs in domain_groups.items():
        if data_source == "ahrefs":
            source_label = "Ahrefs"
            output_key = domain
            domain_label = domain
        else:
            source_label = f"Semrush DB={source_db}"
            db_suffix = "" if source_db == "unknown" else f"_{source_db}"
            output_key = f"{domain}{db_suffix}"
            domain_label = (
                f"{domain} [{source_db}]"
                if source_db != "unknown"
                else domain
            )

        print("\n" + "-" * 68)
        print(f"Target: {domain} | Source: {source_label}")
        print("-" * 68)

        df_source = pd.concat(dfs, ignore_index=True)

        # 空导出也必须保持 URL / Canonical_Key 为字符串类型。
        # pandas 对空列表/空 Series 可能推断为 float64；如果 Sitemap 也为空，
        # merge 会出现 “float64 vs object” 的类型冲突。这里统一固定 dtype。
        if "Source_URL" not in df_source.columns:
            df_source["Source_URL"] = pd.Series(dtype="string")
        else:
            df_source["Source_URL"] = df_source["Source_URL"].astype("string")

        for col, dtype in (
            ("Traffic", "float64"),
            ("Keywords", "float64"),
            ("Top_Keyword", "string"),
        ):
            if col not in df_source.columns:
                df_source[col] = pd.Series(dtype=dtype)

        df_source["Canonical_Key"] = (
            df_source["Source_URL"]
            .map(get_canonical_key)
            .astype("string")
        )
        df_source = (
            df_source.sort_values("Traffic", ascending=False)
            .drop_duplicates("Canonical_Key")
        )

        if domain not in sitemap_cache:
            crawler = SitemapCrawler(domain)
            sitemap_cache[domain] = crawler.run()
        sitemap_urls = sitemap_cache[domain]

        # 显式指定 string dtype：即使 Sitemap 一个 URL 都没抓到，
        # Canonical_Key 也不会被 pandas 推断成 float64。
        df_sitemap = pd.DataFrame({
            "Sitemap_URL": pd.Series(list(sitemap_urls), dtype="string")
        })
        df_sitemap["Canonical_Key"] = (
            df_sitemap["Sitemap_URL"]
            .map(get_canonical_key)
            .astype("string")
        )
        df_sitemap = df_sitemap.drop_duplicates("Canonical_Key")

        merged_df = pd.merge(
            df_sitemap,
            df_source,
            on="Canonical_Key",
            how="outer"
        )
        merged_df["Final_URL"] = merged_df["Source_URL"].combine_first(
            merged_df["Sitemap_URL"]
        )
        merged_df["Traffic"] = merged_df.get("Traffic", 0).fillna(0)
        merged_df["Keywords"] = merged_df.get("Keywords", 0).fillna(0)
        merged_df["Top_Keyword"] = merged_df.get("Top_Keyword", "").fillna("")
        merged_df["In_Sitemap"] = merged_df["Sitemap_URL"].notna()

        logger.info(
            f"Step 1/5: build tree from {len(merged_df):,} URLs"
        )
        tree_data = build_tree_structure(merged_df)

        logger.info("Step 2/5: write TXT tree")
        tree_file = os.path.join(
            WORKSPACE_DIR,
            f"{output_key}_seo_architecture_tree.txt"
        )
        with open(tree_file, "w", encoding="utf-8") as f:
            f.write(f"{domain_label} SEO architecture\n")
            f.write(f"Data source: {source_label}\n")
            f.write("=" * 68 + "\n")
            write_tree_to_txt(tree_data, f)

        logger.info("Step 3/5: write traffic Treemap")
        dashboard_file = os.path.join(
            WORKSPACE_DIR,
            f"{output_key}_seo_dashboard.html"
        )
        generate_dashboard(domain_label, dashboard_file, tree_data)

        logger.info("Step 4/5: write lazy tree chunks")
        asset_dir_name = f"{output_key}_seo_tree_assets"
        asset_dir = os.path.join(WORKSPACE_DIR, asset_dir_name)

        # Dashboard root 已显示域名，因此分片从域名直属子节点开始。
        domain_root = tree_data.get(domain)
        if domain_root is not None:
            lazy_tree_source = domain_root.get("__children__", {})
            root_total_pages = domain_root.get("__total_pages__", 0)
            root_total_traffic = domain_root.get("__total_traffic__", 0)
        else:
            lazy_tree_source = tree_data
            root_total_pages = sum(
                node.get("__total_pages__", 0)
                for node in tree_data.values()
            )
            root_total_traffic = sum(
                node.get("__total_traffic__", 0)
                for node in tree_data.values()
            )

        root_chunk_id, manifest = write_lazy_tree_assets(
            lazy_tree_source,
            asset_dir
        )
        append_lazy_tree_to_dashboard(
            dashboard_file,
            domain_label,
            asset_dir_name,
            root_chunk_id,
            manifest,
            root_total_pages=root_total_pages,
            root_total_traffic=root_total_traffic,
            data_source_label=source_label
        )

        logger.info("Step 5/5: write Excel summary")
        excel_file = os.path.join(
            WORKSPACE_DIR,
            f"{output_key}_seo_metrics.xlsx"
        )

        merged_df["Dir_L1"] = merged_df["Final_URL"].apply(
            lambda x: (
                urlparse(x).path.strip("/").split("/")[0]
                if isinstance(x, str) and urlparse(x).path.strip("/")
                else "Home"
            )
        )
        summary = (
            merged_df.groupby("Dir_L1")
            .agg(
                Pages=("Final_URL", "count"),
                Traffic=("Traffic", "sum")
            )
            .sort_values("Traffic", ascending=False)
        )

        with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
            summary.to_excel(writer, sheet_name="目录流量汇总")

            top_traffic = merged_df[merged_df["Traffic"] > 0].sort_values(
                "Traffic",
                ascending=False
            )
            if top_traffic.empty:
                top_traffic = merged_df.head(1000)

            if len(top_traffic) > 1_048_000:
                logger.warning(
                    f"Traffic-page rows exceed Excel limit: {len(top_traffic):,}; "
                    "writing first 1,048,000 rows."
                )
                top_traffic = top_traffic.head(1_048_000)

            top_traffic.to_excel(
                writer,
                sheet_name="有流量页面清单",
                index=False
            )

        print(f"Done: {domain_label}")
        print(f"TXT: {tree_file}")
        print(f"HTML: {dashboard_file}")
        print(f"Assets: {asset_dir}")
        print(f"Excel: {excel_file}\n")

        del tree_data
        del merged_df
        del df_sitemap
        del df_source


if __name__ == "__main__":
    main()
