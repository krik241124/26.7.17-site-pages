#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pet Lead Classifier — low-cost first-pass for Pet-first vs Pet-adjacent leads.

Workflow:
1) Read domains from .md/.txt/.csv
2) Fetch homepage + robots.txt + sitemap(s) only (no Playwright, no product-page crawl)
3) Score pet relevance + ecommerce signals
4) Output PET_FIRST_AUTO / PET_ADJACENT_AUTO / AMBIGUOUS_REVIEW / REJECT_AUTO

Install:
    pip install requests beautifulsoup4

Run:
    python pet_lead_classifier.py "Pasted markdown(20260903-083623).md" -o pet_leads_stage1.csv

Optional tuning:
    python pet_lead_classifier.py input.md -o out.csv --workers 24 --timeout 8 --max-urls 12000
"""

from __future__ import annotations

import argparse
import csv
import gzip
import io
import re
import sys
import time
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0 Safari/537.36"
)

# Strong pet-commercial vocabulary. Keep this conservative: false positives cost sales time.
PET_TERMS = {
    "pet", "pets", "dog", "dogs", "puppy", "puppies", "canine", "cat", "cats",
    "kitten", "kittens", "feline", "kennel", "kennels", "crate", "crates",
    "leash", "leashes", "harness", "harnesses", "collar", "collars", "cat-tree",
    "cat-trees", "cat_tree", "cat_trees", "scratcher", "scratchers", "litter",
    "dog-bed", "dog-beds", "cat-bed", "cat-beds", "pet-bed", "pet-beds",
    "pet-furniture", "dog-house", "dog-houses", "carrier", "carriers", "pet-stroller",
    "pet-strollers", "pet-bowl", "pet-bowls", "pet-gate", "pet-gates", "paw", "paws",
}

# E-commerce / catalog URL cues.
COMMERCE_PATH_TERMS = {
    "product", "products", "collection", "collections", "shop", "store", "category",
    "categories", "catalog", "catalogue", "buy", "cart", "checkout",
}

# Content/service cues. These do not automatically reject a domain unless commerce signals are weak.
NON_COMMERCE_TERMS = {
    "veterinary", "veterinarian", "animal hospital", "vet clinic", "rescue", "shelter",
    "adoption", "breeder", "breeding", "training school", "dog training", "magazine",
    "news", "forum", "association", "foundation", "charity", "nonprofit", "non-profit",
    "wiki", "encyclopedia", "directory", "insurance", "veterinary clinic",
}

# Obvious giant marketplaces/platforms/retail chains that are not useful SMB lead targets.
# Extend this file over time; it is intentionally small and explicit.
DEFAULT_EXCLUDE_DOMAINS = {
    "amazon.com", "aliexpress.com", "alibaba.com", "temu.com", "ebay.com", "etsy.com",
    "walmart.com", "target.com", "wayfair.com", "chewy.com", "petsmart.com", "petco.com",
    "zazzle.com", "faire.com", "dhgate.com", "made-in-china.com", "kickstarter.com",
    "yelp.com", "pinterest.com", "facebook.com", "instagram.com", "reddit.com",
}

SITEMAP_CANDIDATES = (
    "/sitemap.xml",
    "/sitemap_index.xml",
    "/sitemap-index.xml",
    "/sitemap/sitemap.xml",
)

# Prefer catalog maps; de-prioritize pure content maps.
SITEMAP_PRIORITY_GOOD = ("product", "collection", "category", "shop", "catalog")
SITEMAP_PRIORITY_BAD = ("blog", "post", "article", "news", "author", "tag")


@dataclass
class Result:
    domain: str
    final_url: str = ""
    http_status: str = ""
    title: str = ""
    meta_description: str = ""
    homepage_pet_hits: int = 0
    homepage_ecom_hits: int = 0
    homepage_noncommerce_hits: int = 0
    sitemap_found: int = 0
    sitemap_files_scanned: int = 0
    sitemap_urls_scanned: int = 0
    pet_urls: int = 0
    commerce_urls: int = 0
    pet_commerce_urls: int = 0
    pet_url_ratio: float = 0.0
    pet_commerce_ratio: float = 0.0
    score_pet: float = 0.0
    score_ecom: float = 0.0
    classification: str = "AMBIGUOUS_REVIEW"
    confidence: str = "low"
    reason: str = ""
    error: str = ""
    seconds: float = 0.0


def normalize_domain(raw: str) -> str:
    s = raw.strip().lower().replace("\\.", ".")
    s = re.sub(r"^https?://", "", s)
    s = s.split("/")[0].strip(".")
    if s.startswith("www."):
        s = s[4:]
    return s


def read_domains(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8", errors="ignore")
    out: list[str] = []
    seen = set()
    for line in text.splitlines():
        candidates = []
        if "|" in line:  # markdown table
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            candidates.extend(cells)
        else:
            candidates.extend(re.split(r"[,\t; ]+", line.strip()))
        for c in candidates:
            d = normalize_domain(c)
            if not d or d in {"domain"} or set(d) <= {"-", ":"}:
                continue
            if "." not in d or " " in d:
                continue
            if re.fullmatch(r"[a-z0-9.-]+", d) and d not in seen:
                seen.add(d)
                out.append(d)
    return out


def tokens_from_text(text: str) -> list[str]:
    t = text.lower()
    t = re.sub(r"https?://", " ", t)
    t = t.replace("_", "-")
    return re.findall(r"[a-z0-9]+(?:-[a-z0-9]+)*", t)


def count_pet_terms(text: str) -> int:
    t = text.lower().replace("_", "-")
    count = 0
    for term in PET_TERMS:
        if "-" in term or " " in term:
            count += t.count(term)
        else:
            count += len(re.findall(rf"\b{re.escape(term)}\b", t))
    return count


def count_terms(text: str, terms: Iterable[str]) -> int:
    t = text.lower()
    return sum(t.count(x) for x in terms)


def same_domain(url: str, domain: str) -> bool:
    host = (urlparse(url).hostname or "").lower()
    return host == domain or host == "www." + domain or host.endswith("." + domain)


def session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    })
    return s


def get(session_: requests.Session, url: str, timeout: float, max_bytes: int = 4_000_000):
    r = session_.get(url, timeout=timeout, allow_redirects=True, stream=True)
    content = bytearray()
    for chunk in r.iter_content(65536):
        if chunk:
            content.extend(chunk)
            if len(content) >= max_bytes:
                break
    return r, bytes(content)


def fetch_homepage(s: requests.Session, domain: str, timeout: float):
    last_err = None
    for scheme in ("https://", "http://"):
        try:
            r, raw = get(s, scheme + domain + "/", timeout, max_bytes=2_000_000)
            if r.status_code < 500:
                return r, raw
        except Exception as e:
            last_err = e
    if last_err:
        raise last_err
    raise RuntimeError("homepage unavailable")


def homepage_signals(raw: bytes):
    text = raw.decode("utf-8", errors="ignore")
    soup = BeautifulSoup(text, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    meta = ""
    m = soup.find("meta", attrs={"name": re.compile("description", re.I)})
    if m and m.get("content"):
        meta = str(m.get("content")).strip()

    # Heavier weight to title/meta/nav/headings, lighter to all visible text.
    nav_parts = []
    for tag in soup.find_all(["nav", "header", "h1", "h2", "a"]):
        txt = tag.get_text(" ", strip=True)
        if txt:
            nav_parts.append(txt)
        if len(nav_parts) >= 500:
            break
    focused = " ".join([title, meta] + nav_parts)
    visible = soup.get_text(" ", strip=True)[:250_000]

    pet_hits = min(50, count_pet_terms(focused) * 2 + count_pet_terms(visible) // 10)

    ecom_needles = [
        "add to cart", "add to bag", "shopping cart", "checkout", "shop now", "buy now",
        "product", "products", "collections", "shipping", "returns", "price", "currency",
        "application/ld+json", '"@type":"product"', '"@type": "product"',
    ]
    ecom_hits = min(50, count_terms(text, ecom_needles))
    non_hits = min(50, count_terms((title + " " + meta + " " + visible[:100_000]), NON_COMMERCE_TERMS))
    return title[:250], meta[:500], pet_hits, ecom_hits, non_hits


def extract_sitemaps_from_robots(s: requests.Session, base_url: str, timeout: float) -> list[str]:
    try:
        robots = urljoin(base_url, "/robots.txt")
        r, raw = get(s, robots, timeout, max_bytes=500_000)
        if r.status_code >= 400:
            return []
        txt = raw.decode("utf-8", errors="ignore")
        out = []
        for line in txt.splitlines():
            if line.lower().startswith("sitemap:"):
                u = line.split(":", 1)[1].strip()
                if u.startswith("http"):
                    out.append(u)
        return out[:20]
    except Exception:
        return []


def parse_xml_locs(raw: bytes):
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    root = ET.fromstring(raw)
    tag = root.tag.lower()
    locs = []
    for elem in root.iter():
        if elem.tag.lower().endswith("loc") and elem.text:
            locs.append(elem.text.strip())
    is_index = tag.endswith("sitemapindex")
    return is_index, locs


def sitemap_priority(url: str) -> tuple[int, str]:
    x = url.lower()
    score = 0
    if any(k in x for k in SITEMAP_PRIORITY_GOOD):
        score += 10
    if any(k in x for k in SITEMAP_PRIORITY_BAD):
        score -= 8
    return (-score, x)


def scan_sitemaps(s: requests.Session, domain: str, base_url: str, timeout: float,
                  max_urls: int, max_sitemaps: int):
    seeds = extract_sitemaps_from_robots(s, base_url, timeout)
    if not seeds:
        seeds = [urljoin(base_url, x) for x in SITEMAP_CANDIDATES]

    queue = list(dict.fromkeys(seeds))
    seen_maps = set()
    urls_seen = set()
    sitemap_found = False

    while queue and len(seen_maps) < max_sitemaps and len(urls_seen) < max_urls:
        queue.sort(key=sitemap_priority)
        sm = queue.pop(0)
        if sm in seen_maps:
            continue
        seen_maps.add(sm)
        try:
            r, raw = get(s, sm, timeout, max_bytes=5_000_000)
            if r.status_code >= 400 or not raw:
                continue
            ctype = (r.headers.get("content-type") or "").lower()
            if "xml" not in ctype and b"<urlset" not in raw[:5000].lower() and b"<sitemapindex" not in raw[:5000].lower():
                continue
            is_index, locs = parse_xml_locs(raw)
            sitemap_found = True
            if is_index:
                # Catalog maps first, content maps last. Cap aggressively for low cost.
                children = [u for u in locs if u.startswith("http") and same_domain(u, domain)]
                children.sort(key=sitemap_priority)
                for u in children:
                    if u not in seen_maps and u not in queue:
                        queue.append(u)
                queue = queue[: max_sitemaps * 3]
            else:
                for u in locs:
                    if len(urls_seen) >= max_urls:
                        break
                    if u.startswith("http") and same_domain(u, domain):
                        urls_seen.add(u)
        except Exception:
            continue

    pet_urls = 0
    commerce_urls = 0
    pet_commerce_urls = 0
    for u in urls_seen:
        p = urlparse(u).path.lower().replace("_", "-")
        pet = count_pet_terms(p) > 0
        commerce = any(f"/{x}" in p or p.startswith(f"/{x}") for x in COMMERCE_PATH_TERMS)
        # Product/category slugs themselves often carry strong pet nouns even if /products/ is absent.
        if pet:
            pet_urls += 1
        if commerce:
            commerce_urls += 1
        if pet and commerce:
            pet_commerce_urls += 1

    total = len(urls_seen)
    return {
        "sitemap_found": int(sitemap_found),
        "sitemap_files_scanned": len(seen_maps),
        "sitemap_urls_scanned": total,
        "pet_urls": pet_urls,
        "commerce_urls": commerce_urls,
        "pet_commerce_urls": pet_commerce_urls,
        "pet_url_ratio": pet_urls / total if total else 0.0,
        "pet_commerce_ratio": pet_commerce_urls / max(commerce_urls, 1),
    }


def classify(r: Result) -> Result:
    d = r.domain
    if d in DEFAULT_EXCLUDE_DOMAINS:
        r.classification = "REJECT_AUTO"
        r.confidence = "high"
        r.reason = "explicit giant marketplace/chain/platform exclusion"
        return r

    # Scoring uses multiple independent signals. Sitemap ratios matter most.
    pet_score = 0.0
    pet_score += min(r.homepage_pet_hits, 20) * 0.8
    pet_score += min(r.pet_urls, 80) * 0.15
    pet_score += min(r.pet_commerce_urls, 50) * 0.35
    pet_score += min(r.pet_url_ratio, 1.0) * 18
    pet_score += min(r.pet_commerce_ratio, 1.0) * 22

    ecom_score = min(r.homepage_ecom_hits, 20) * 0.7
    ecom_score += min(r.commerce_urls, 100) * 0.08

    r.score_pet = round(pet_score, 2)
    r.score_ecom = round(ecom_score, 2)

    # Hard non-lead only when commerce evidence is weak.
    if r.homepage_noncommerce_hits >= 3 and ecom_score < 4 and r.pet_commerce_urls < 3:
        r.classification = "REJECT_AUTO"
        r.confidence = "medium"
        r.reason = "service/content/non-commerce signals dominate; weak ecommerce evidence"
        return r

    # High-confidence Pet-first: pet dominates either commerce taxonomy or whole sitemap,
    # and the site has enough commercial evidence.
    if (
        ecom_score >= 4
        and (
            (r.pet_commerce_urls >= 12 and r.pet_commerce_ratio >= 0.55)
            or (r.pet_urls >= 20 and r.pet_url_ratio >= 0.45 and r.homepage_pet_hits >= 4)
            or (r.homepage_pet_hits >= 12 and r.pet_commerce_urls >= 5)
        )
    ):
        r.classification = "PET_FIRST_AUTO"
        r.confidence = "high" if r.pet_commerce_ratio >= 0.65 or r.homepage_pet_hits >= 15 else "medium"
        r.reason = "pet category dominates homepage/catalog signals"
        return r

    # Pet-adjacent: real pet commercial cluster exists, but pet does not dominate the whole business.
    if ecom_score >= 4 and (
        r.pet_commerce_urls >= 5
        or (r.pet_urls >= 8 and r.homepage_pet_hits >= 2)
        or (r.pet_commerce_ratio >= 0.08 and r.commerce_urls >= 15)
    ):
        if r.pet_commerce_ratio < 0.55 or r.homepage_pet_hits < 10:
            r.classification = "PET_ADJACENT_AUTO"
            r.confidence = "medium"
            r.reason = "meaningful pet product/category cluster inside a broader ecommerce site"
            return r

    # Very weak pet evidence: safe auto-reject if we did get enough sitemap/homepage evidence.
    evidence_volume = r.sitemap_urls_scanned >= 20 or r.homepage_ecom_hits >= 3
    if evidence_volume and r.pet_urls == 0 and r.homepage_pet_hits == 0:
        r.classification = "REJECT_AUTO"
        r.confidence = "high"
        r.reason = "no pet signal found in homepage or sitemap sample"
        return r

    # Everything else goes to cheap web/manual verification.
    r.classification = "AMBIGUOUS_REVIEW"
    r.confidence = "low"
    if not r.sitemap_found:
        r.reason = "sitemap unavailable/blocked or insufficient taxonomy evidence"
    elif r.score_ecom < 4:
        r.reason = "pet relevance exists but ecommerce/merchant status is unclear"
    else:
        r.reason = "mixed pet/general signals; needs web verification"
    return r


def analyze_domain(domain: str, timeout: float, max_urls: int, max_sitemaps: int) -> Result:
    t0 = time.time()
    r = Result(domain=domain)
    s = session()
    try:
        resp, raw = fetch_homepage(s, domain, timeout)
        r.final_url = resp.url
        r.http_status = str(resp.status_code)
        title, meta, ph, eh, nh = homepage_signals(raw)
        r.title, r.meta_description = title, meta
        r.homepage_pet_hits, r.homepage_ecom_hits, r.homepage_noncommerce_hits = ph, eh, nh
        base = f"{urlparse(resp.url).scheme}://{urlparse(resp.url).netloc}"
        sm = scan_sitemaps(s, domain, base, timeout, max_urls, max_sitemaps)
        for k, v in sm.items():
            setattr(r, k, v)
    except Exception as e:
        r.error = f"{type(e).__name__}: {str(e)[:220]}"
    r = classify(r)
    r.seconds = round(time.time() - t0, 2)
    return r


def write_csv(path: Path, results: list[Result]):
    fields = list(asdict(Result(domain="")).keys())
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in results:
            row = asdict(r)
            row["pet_url_ratio"] = f"{r.pet_url_ratio:.4f}"
            row["pet_commerce_ratio"] = f"{r.pet_commerce_ratio:.4f}"
            w.writerow(row)


def main():
    ap = argparse.ArgumentParser(description="Low-cost Pet-first / Pet-adjacent lead classifier")
    ap.add_argument("input", type=Path, help="Input .md/.txt/.csv containing domains")
    ap.add_argument("-o", "--output", type=Path, default=Path("pet_leads_stage1.csv"))
    ap.add_argument("--workers", type=int, default=20, help="Concurrent domains (default: 20)")
    ap.add_argument("--timeout", type=float, default=8.0, help="Per-request timeout seconds (default: 8)")
    ap.add_argument("--max-urls", type=int, default=12000, help="Max sitemap URLs sampled/domain")
    ap.add_argument("--max-sitemaps", type=int, default=24, help="Max sitemap files/domain")
    args = ap.parse_args()

    domains = read_domains(args.input)
    if not domains:
        print("No domains found", file=sys.stderr)
        sys.exit(2)

    print(f"Loaded {len(domains)} unique domains")
    results: list[Result] = []
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        futs = {
            ex.submit(analyze_domain, d, args.timeout, args.max_urls, args.max_sitemaps): d
            for d in domains
        }
        total = len(futs)
        for i, fut in enumerate(as_completed(futs), 1):
            d = futs[fut]
            try:
                res = fut.result()
            except Exception as e:
                res = Result(domain=d, error=f"worker: {e}")
            results.append(res)
            if i % 25 == 0 or i == total:
                counts = {}
                for x in results:
                    counts[x.classification] = counts.get(x.classification, 0) + 1
                print(f"[{i}/{total}] {counts}", flush=True)

    order = {d: i for i, d in enumerate(domains)}
    results.sort(key=lambda x: order.get(x.domain, 10**9))
    write_csv(args.output, results)

    # Write just the ambiguous domains for ChatGPT/web verification.
    ambiguous_path = args.output.with_name(args.output.stem + "_ambiguous.txt")
    with ambiguous_path.open("w", encoding="utf-8") as f:
        for x in results:
            if x.classification == "AMBIGUOUS_REVIEW":
                f.write(x.domain + "\n")

    counts = {}
    for x in results:
        counts[x.classification] = counts.get(x.classification, 0) + 1
    print("\nDone")
    print("Output:", args.output)
    print("Ambiguous:", ambiguous_path)
    print("Counts:", counts)


if __name__ == "__main__":
    main()
