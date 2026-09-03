#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pet Lead Classifier V3 — resilient sitemap-first pass for Pet-first vs Pet-adjacent leads.

Workflow:
1) Read domains from .md/.txt/.csv
2) Fetch homepage + robots.txt + sitemap(s) only (no Playwright, no product-page crawl)
3) Score pet relevance + ecommerce signals
4) Output PET_FIRST_AUTO / PET_ADJACENT_AUTO / AMBIGUOUS_REVIEW / REJECT_AUTO

V3 reliability improvements:
- Live progress bar with throughput, elapsed time, ETA and estimated finish clock time
- Checkpoint results to the requested output CSV every N completed domains
- Checkpoint ambiguous-domain TXT at the same time
- Optional --resume from an existing output CSV
- Per-domain wall-clock budget to prevent one pathological sitemap from blocking a whole batch
- Sitemap discovery mirrors the proven page_sitemap crawler: robots.txt Sitemap declarations first, then standard fallbacks
- Homepage failure/block no longer prevents sitemap crawling
- robots-declared and conventional sitemap paths are tried together instead of mutually exclusively
- apex/www + HTTPS/HTTP fallbacks for broken redirects and odd host setups
- SSL verification fallback for legacy/misconfigured small ecommerce sites
- Streaming <loc> extraction: no 5 MB XML truncation and no full-tree ElementTree memory spike
- .xml.gz is detected/decompressed before format validation
- Tolerant sitemap parsing accepts XML served with wrong Content-Type and plain-text sitemap files
- Light HTTP retry for transient 408/425/5xx errors (not repeated hammering on 403/429)
- Properly close streamed HTTP responses
- Known excluded giant platforms are rejected before any network request
- Ctrl+C writes a final checkpoint before stopping submission/waiting logic

Install:
    pip install requests beautifulsoup4

Run:
    python pet_lead_classifier_v3.py data.txt -o pet_leads_stage1.csv

Recommended:
    python pet_lead_classifier_v3.py data.txt -o pet_leads_stage1.csv \
        --workers 20 --timeout 8 --domain-budget 45 --checkpoint-every 25

Resume an interrupted run:
    python pet_lead_classifier_v3.py data.txt -o pet_leads_stage1.csv --resume
"""

from __future__ import annotations

import argparse
import csv
import gzip
import html
import zlib
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin, urlparse

import requests
import urllib3
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

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
DEFAULT_EXCLUDE_DOMAINS = {
    "amazon.com", "aliexpress.com", "alibaba.com", "temu.com", "ebay.com", "etsy.com",
    "walmart.com", "target.com", "wayfair.com", "chewy.com", "petsmart.com", "petco.com",
    "zazzle.com", "faire.com", "dhgate.com", "made-in-china.com", "kickstarter.com",
    "yelp.com", "pinterest.com", "facebook.com", "instagram.com", "reddit.com",
    "unsplash.com",
}

SITEMAP_CANDIDATES = (
    "/sitemap.xml",
    "/sitemap_index.xml",
    "/sitemap-index.xml",
    "/wp-sitemap.xml",
    "/sitemap/sitemap.xml",
    "/sitemap.xml.gz",
    "/sitemap_index.xml.gz",
)

# Prefer catalog maps; de-prioritize pure content maps.
SITEMAP_PRIORITY_GOOD = ("product", "products", "collection", "category", "product_cat", "shop", "store", "catalog", "woocommerce")
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
    homepage_blocked: int = 0
    robots_status: str = ""
    robots_sitemaps: int = 0
    sitemap_found: int = 0
    sitemap_files_scanned: int = 0
    sitemap_urls_scanned: int = 0
    sitemap_http_errors: int = 0
    sitemap_parse_errors: int = 0
    sitemap_gzip_files: int = 0
    sitemap_discovery: str = ""
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
    domain_budget_exceeded: int = 0


class DomainBudgetExceeded(TimeoutError):
    """Raised/recorded when one domain exceeds its wall-clock analysis budget."""


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
        "Accept": "text/html,application/xhtml+xml,application/xml,text/xml,text/plain;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Cache-Control": "no-cache",
    })
    # Retry transient infrastructure errors once. Do NOT retry 403/429 aggressively:
    # that usually makes anti-bot handling worse and wastes the per-domain budget.
    retry = Retry(
        total=1,
        connect=1,
        read=1,
        status=1,
        backoff_factor=0.25,
        status_forcelist=(408, 425, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "HEAD"}),
        raise_on_status=False,
        respect_retry_after_header=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=8, pool_maxsize=8)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


def _remaining_request_timeout(timeout: float, deadline: float | None) -> float:
    if deadline is None:
        return timeout
    remaining = deadline - time.monotonic()
    if remaining <= 0:
        raise DomainBudgetExceeded("domain wall-clock budget exceeded")
    return max(0.25, min(timeout, remaining))


def _request(
    session_: requests.Session,
    url: str,
    timeout: float,
    deadline: float | None,
    *,
    stream: bool,
):
    """GET with a one-time SSL-verification fallback for public crawl-only reads."""
    req_timeout = _remaining_request_timeout(timeout, deadline)
    kwargs = dict(timeout=req_timeout, allow_redirects=True, stream=stream)
    try:
        return session_.get(url, verify=True, **kwargs)
    except requests.exceptions.SSLError:
        req_timeout = _remaining_request_timeout(timeout, deadline)
        kwargs["timeout"] = req_timeout
        return session_.get(url, verify=False, **kwargs)


def get(
    session_: requests.Session,
    url: str,
    timeout: float,
    max_bytes: int = 4_000_000,
    deadline: float | None = None,
):
    """GET a bounded amount of non-sitemap content and always close the response."""
    r = _request(session_, url, timeout, deadline, stream=True)
    content = bytearray()
    try:
        for chunk in r.iter_content(65536):
            if deadline is not None and time.monotonic() >= deadline:
                raise DomainBudgetExceeded("domain wall-clock budget exceeded while streaming")
            if not chunk:
                continue
            remaining_bytes = max_bytes - len(content)
            if remaining_bytes <= 0:
                break
            content.extend(chunk[:remaining_bytes])
            if len(content) >= max_bytes:
                break
        return r, bytes(content)
    finally:
        r.close()


def fetch_homepage(s: requests.Session, domain: str, timeout: float, deadline: float | None):
    last_err = None
    # Try the most common canonical forms. A 403/429 is still returned as evidence,
    # but sitemap crawling continues independently afterward.
    candidates = (
        f"https://{domain}/",
        f"https://www.{domain}/",
        f"http://{domain}/",
        f"http://www.{domain}/",
    )
    best = None
    for url in candidates:
        try:
            r, raw = get(s, url, timeout, max_bytes=2_000_000, deadline=deadline)
            if best is None:
                best = (r, raw)
            if 200 <= r.status_code < 400:
                return r, raw
            # A block page is useful for status/diagnostics; don't burn the entire
            # budget trying every variant after we already know the site is alive.
            if r.status_code in (401, 403, 429):
                return r, raw
        except DomainBudgetExceeded:
            raise
        except Exception as e:
            last_err = e
    if best is not None:
        return best
    if last_err:
        raise last_err
    raise RuntimeError("homepage unavailable")


def looks_like_block_page(status: int, raw: bytes) -> bool:
    if status in (401, 403, 429):
        return True
    sample = raw[:200_000].decode("utf-8", errors="ignore").lower()
    needles = (
        "just a moment", "verifying your connection", "checking your browser",
        "attention required", "cf-chl-", "cloudflare ray id", "access denied",
        "enable javascript and cookies to continue",
    )
    return any(x in sample for x in needles)


def homepage_signals(raw: bytes):
    text = raw.decode("utf-8", errors="ignore")
    soup = BeautifulSoup(text, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    meta = ""
    m = soup.find("meta", attrs={"name": re.compile("description", re.I)})
    if m and m.get("content"):
        meta = str(m.get("content")).strip()

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


def _base_variants(domain: str, final_url: str = "") -> list[str]:
    out: list[str] = []
    if final_url:
        p = urlparse(final_url)
        if p.scheme and p.netloc:
            out.append(f"{p.scheme}://{p.netloc}")
    out.extend([
        f"https://{domain}",
        f"https://www.{domain}",
        f"http://{domain}",
        f"http://www.{domain}",
    ])
    return list(dict.fromkeys(x.rstrip("/") for x in out))


def extract_sitemaps_from_robots(
    s: requests.Session,
    bases: list[str],
    timeout: float,
    deadline: float | None,
) -> tuple[list[str], str]:
    """Read robots from preferred host first; try alternate host only when needed."""
    statuses: list[str] = []
    collected: list[str] = []

    for i, base_url in enumerate(bases[:3]):
        try:
            robots = urljoin(base_url + "/", "robots.txt")
            r, raw = get(s, robots, timeout, max_bytes=1_000_000, deadline=deadline)
            statuses.append(f"{urlparse(base_url).netloc}:{r.status_code}")
            if r.status_code >= 400:
                # Try another host variant if this one is blocked/missing.
                continue
            txt = raw.decode("utf-8", errors="ignore")
            for m in re.finditer(r"(?im)^\s*Sitemap\s*:\s*(\S+)\s*$", txt):
                u = html.unescape(m.group(1).strip())
                if u.startswith(("http://", "https://")):
                    collected.append(u)
            # A successfully fetched robots.txt is authoritative enough; if it has
            # no Sitemap directives we'll rely on conventional paths on this base.
            if r.status_code == 200:
                break
        except DomainBudgetExceeded:
            raise
        except Exception as e:
            statuses.append(f"{urlparse(base_url).netloc}:{type(e).__name__}")
            continue

    return list(dict.fromkeys(collected))[:50], ";".join(statuses)


def _decode_loc(raw: bytes) -> str:
    text = raw.decode("utf-8", errors="ignore").strip()
    if text.startswith("<![CDATA[") and text.endswith("]]>"):
        text = text[9:-3].strip()
    return html.unescape(text)


def _stream_sitemap(
    s: requests.Session,
    sitemap_url: str,
    timeout: float,
    deadline: float | None,
    max_locs: int,
    max_uncompressed_bytes: int = 60_000_000,
):
    """Stream a sitemap and extract <loc> values without building a full XML tree.

    This intentionally behaves more like the proven page_sitemap crawler's tolerant
    <loc> extraction than strict ElementTree.fromstring(). It can stop early after
    enough URLs are sampled, while still supporting .xml.gz.
    """
    r = _request(s, sitemap_url, timeout, deadline, stream=True)
    locs: list[str] = []
    kind = "unknown"  # index / urlset / text / unknown
    gzip_body = False
    parse_error = False
    decompressed_total = 0
    buffer = bytearray()
    plain = bytearray()
    root_probe = bytearray()
    decomp = None
    loc_re = re.compile(rb"<loc\b[^>]*>\s*(.*?)\s*</loc\s*>", re.I | re.S)

    try:
        status = r.status_code
        ctype = (r.headers.get("content-type") or "").lower()
        if status >= 400:
            return status, kind, locs, gzip_body, True

        first_data_seen = False
        for chunk in r.iter_content(65536):
            if deadline is not None and time.monotonic() >= deadline:
                raise DomainBudgetExceeded("domain wall-clock budget exceeded while reading sitemap")
            if not chunk:
                continue

            if not first_data_seen:
                first_data_seen = True
                if chunk[:2] == b"\x1f\x8b":
                    gzip_body = True
                    decomp = zlib.decompressobj(16 + zlib.MAX_WBITS)

            try:
                data = decomp.decompress(chunk) if decomp else chunk
            except zlib.error:
                parse_error = True
                break
            if not data:
                continue

            decompressed_total += len(data)
            if decompressed_total > max_uncompressed_bytes:
                # Safety valve for malformed/no-loc giant responses. Normal large
                # sitemaps should hit max_locs and stop long before this.
                parse_error = True
                break

            if len(root_probe) < 131072:
                root_probe.extend(data[: 131072 - len(root_probe)])
                low = bytes(root_probe).lower()
                if b"<sitemapindex" in low:
                    kind = "index"
                elif b"<urlset" in low:
                    kind = "urlset"

            # Keep a bounded copy only for possible plain-text sitemap fallback.
            if kind == "unknown" and len(plain) < 8_000_000:
                remain = 8_000_000 - len(plain)
                plain.extend(data[:remain])

            buffer.extend(data)
            last_end = 0
            for m in loc_re.finditer(buffer):
                u = _decode_loc(m.group(1))
                if u:
                    locs.append(u)
                last_end = m.end()
                target = max_locs
                if kind == "index":
                    target = min(max_locs, 500)
                if len(locs) >= max(1, target):
                    return status, kind, locs, gzip_body, parse_error

            if last_end:
                del buffer[:last_end]
            elif len(buffer) > 262144:
                # Preserve enough tail for a <loc> split across chunks.
                del buffer[:-32768]

        if decomp:
            try:
                tail = decomp.flush()
                if tail:
                    buffer.extend(tail)
                    for m in loc_re.finditer(buffer):
                        u = _decode_loc(m.group(1))
                        if u:
                            locs.append(u)
                            if len(locs) >= max_locs:
                                break
            except zlib.error:
                parse_error = True

        # Accept Google-compatible plain text sitemaps too.
        if not locs and kind == "unknown" and (
            "text/plain" in ctype or sitemap_url.lower().endswith(".txt")
        ):
            for line in plain.decode("utf-8", errors="ignore").splitlines():
                u = html.unescape(line.strip())
                if u.startswith(("http://", "https://")):
                    locs.append(u)
                    if len(locs) >= max_locs:
                        break
            if locs:
                kind = "text"

        if kind == "unknown" and locs:
            # Broken XML can omit/garble root tags while still exposing valid locs.
            kind = "urlset"
        if not locs:
            parse_error = True
        return status, kind, locs, gzip_body, parse_error
    finally:
        r.close()


def sitemap_priority(url: str) -> tuple[int, str]:
    x = url.lower()
    score = 0
    if any(k in x for k in SITEMAP_PRIORITY_GOOD):
        score += 30
    if "sitemap" in x and ("index" in x or x.rstrip("/").endswith("sitemap.xml")):
        score += 8
    if any(k in x for k in SITEMAP_PRIORITY_BAD):
        score -= 12
    return (-score, x)


def scan_sitemaps(
    s: requests.Session,
    domain: str,
    final_url: str,
    timeout: float,
    max_urls: int,
    max_sitemaps: int,
    deadline: float | None,
):
    budget_exceeded = False
    bases = _base_variants(domain, final_url)
    primary = bases[0] if bases else f"https://{domain}"

    try:
        robots_seeds, robots_status = extract_sitemaps_from_robots(s, bases, timeout, deadline)
    except DomainBudgetExceeded:
        robots_seeds, robots_status = [], "budget"
        budget_exceeded = True

    # Important V3 change: robots declarations and conventional paths are UNIONED.
    # A stale robots Sitemap line should not suppress /sitemap.xml probing.
    standard_primary = [urljoin(primary + "/", p.lstrip("/")) for p in SITEMAP_CANDIDATES]

    # Cheap alternate-host recovery: only the three most common roots on up to two
    # alternate bases. This catches sites where homepage/robots canonicalization is odd.
    standard_alt: list[str] = []
    for b in bases[1:3]:
        for p in ("/sitemap.xml", "/sitemap_index.xml", "/wp-sitemap.xml"):
            standard_alt.append(urljoin(b + "/", p.lstrip("/")))

    fallback_seeds = list(dict.fromkeys(standard_primary + standard_alt))
    # Honor robots declarations first, then conventional probes in priority order.
    # Child maps discovered from an index will be inserted at the FRONT later.
    seeds = list(dict.fromkeys(robots_seeds + sorted(fallback_seeds, key=sitemap_priority)))
    queue = seeds[:]
    seen_maps: set[str] = set()
    urls_seen: set[str] = set()
    sitemap_found = False
    http_errors = 0
    parse_errors = 0
    gzip_files = 0
    discovery_parts: list[str] = []

    while queue and len(seen_maps) < max_sitemaps and len(urls_seen) < max_urls:
        if deadline is not None and time.monotonic() >= deadline:
            budget_exceeded = True
            break

        sm = queue.pop(0)
        if sm in seen_maps:
            continue
        seen_maps.add(sm)

        try:
            remaining_urls = max(1, max_urls - len(urls_seen))
            # Indexes need only enough child sitemap URLs to fill our map budget.
            max_locs = max(remaining_urls, max_sitemaps * 4)
            status, kind, locs, was_gzip, parse_bad = _stream_sitemap(
                s, sm, timeout, deadline, max_locs=max_locs
            )
            if was_gzip:
                gzip_files += 1
            if status >= 400:
                http_errors += 1
                continue
            if parse_bad and not locs:
                parse_errors += 1
                continue
            if not locs:
                continue

            sitemap_found = True
            if sm in robots_seeds and "robots" not in discovery_parts:
                discovery_parts.append("robots")
            if sm in standard_primary and "standard" not in discovery_parts:
                discovery_parts.append("standard")
            if sm in standard_alt and "alternate-host" not in discovery_parts:
                discovery_parts.append("alternate-host")

            if kind == "index":
                children = [
                    u for u in locs
                    if u.startswith(("http://", "https://")) and same_domain(u, domain)
                ]
                children.sort(key=sitemap_priority)
                new_children = [
                    u for u in children
                    if u not in seen_maps and u not in queue
                ]
                # Critical: recurse into a discovered sitemap index BEFORE burning
                # the map budget on guessed fallback paths such as /wp-sitemap.xml.
                queue = new_children + queue
                queue = queue[: max_sitemaps * 6]
            else:
                for u in locs:
                    if len(urls_seen) >= max_urls:
                        break
                    if u.startswith(("http://", "https://")) and same_domain(u, domain):
                        urls_seen.add(u)
        except DomainBudgetExceeded:
            budget_exceeded = True
            break
        except Exception:
            parse_errors += 1
            continue

    pet_urls = 0
    commerce_urls = 0
    pet_commerce_urls = 0
    for u in urls_seen:
        p = urlparse(u).path.lower().replace("_", "-")
        pet = count_pet_terms(p) > 0
        commerce = any(f"/{x}" in p or p.startswith(f"/{x}") for x in COMMERCE_PATH_TERMS)
        if pet:
            pet_urls += 1
        if commerce:
            commerce_urls += 1
        if pet and commerce:
            pet_commerce_urls += 1

    total = len(urls_seen)
    return {
        "robots_status": robots_status,
        "robots_sitemaps": len(robots_seeds),
        "sitemap_found": int(sitemap_found),
        "sitemap_files_scanned": len(seen_maps),
        "sitemap_urls_scanned": total,
        "sitemap_http_errors": http_errors,
        "sitemap_parse_errors": parse_errors,
        "sitemap_gzip_files": gzip_files,
        "sitemap_discovery": "+".join(discovery_parts),
        "pet_urls": pet_urls,
        "commerce_urls": commerce_urls,
        "pet_commerce_urls": pet_commerce_urls,
        "pet_url_ratio": pet_urls / total if total else 0.0,
        "pet_commerce_ratio": pet_commerce_urls / max(commerce_urls, 1),
        "domain_budget_exceeded": int(budget_exceeded),
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
    if r.domain_budget_exceeded:
        r.reason = "domain time budget exceeded; partial evidence only"
    elif not r.sitemap_found:
        r.reason = "sitemap unavailable/blocked or insufficient taxonomy evidence"
    elif r.score_ecom < 4:
        r.reason = "pet relevance exists but ecommerce/merchant status is unclear"
    else:
        r.reason = "mixed pet/general signals; needs web verification"
    return r


def analyze_domain(
    domain: str,
    timeout: float,
    max_urls: int,
    max_sitemaps: int,
    domain_budget: float,
) -> Result:
    t0 = time.monotonic()
    r = Result(domain=domain)

    if domain in DEFAULT_EXCLUDE_DOMAINS:
        r = classify(r)
        r.seconds = round(time.monotonic() - t0, 2)
        return r

    deadline = t0 + domain_budget if domain_budget > 0 else None
    s = session()
    final_url = ""
    homepage_error = ""
    sitemap_error = ""

    # Homepage and sitemap are deliberately independent in V3.
    try:
        resp, raw = fetch_homepage(s, domain, timeout, deadline)
        final_url = resp.url
        r.final_url = resp.url
        r.http_status = str(resp.status_code)
        r.homepage_blocked = int(looks_like_block_page(resp.status_code, raw))
        if not r.homepage_blocked:
            title, meta, ph, eh, nh = homepage_signals(raw)
            r.title, r.meta_description = title, meta
            r.homepage_pet_hits, r.homepage_ecom_hits, r.homepage_noncommerce_hits = ph, eh, nh
        else:
            # Keep a minimal title for diagnostics, but don't score a challenge page.
            try:
                title, meta, _, _, _ = homepage_signals(raw)
                r.title, r.meta_description = title, meta
            except Exception:
                pass
    except DomainBudgetExceeded as e:
        r.domain_budget_exceeded = 1
        homepage_error = f"homepage budget: {str(e)[:120]}"
    except Exception as e:
        homepage_error = f"homepage {type(e).__name__}: {str(e)[:160]}"

    # Even when the homepage is blocked/unavailable, sitemap discovery still runs.
    if not r.domain_budget_exceeded:
        try:
            sm = scan_sitemaps(
                s, domain, final_url, timeout, max_urls, max_sitemaps, deadline
            )
            for k, v in sm.items():
                setattr(r, k, v)
        except DomainBudgetExceeded as e:
            r.domain_budget_exceeded = 1
            sitemap_error = f"sitemap budget: {str(e)[:120]}"
        except Exception as e:
            sitemap_error = f"sitemap {type(e).__name__}: {str(e)[:160]}"

    s.close()
    r.error = " | ".join(x for x in (homepage_error, sitemap_error) if x)
    r = classify(r)
    r.seconds = round(time.monotonic() - t0, 2)
    return r

def _row_from_result(r: Result) -> dict:
    row = asdict(r)
    row["pet_url_ratio"] = f"{r.pet_url_ratio:.4f}"
    row["pet_commerce_ratio"] = f"{r.pet_commerce_ratio:.4f}"
    return row


def write_csv(path: Path, results: list[Result]):
    """Atomically replace the output CSV so every checkpoint is a complete valid file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = list(asdict(Result(domain="")).keys())
    tmp = path.with_name(path.name + ".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in results:
            w.writerow(_row_from_result(r))
        f.flush()
    tmp.replace(path)


def write_ambiguous(path: Path, results: list[Result]):
    tmp = path.with_name(path.name + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        for x in results:
            if x.classification == "AMBIGUOUS_REVIEW":
                f.write(x.domain + "\n")
        f.flush()
    tmp.replace(path)


def checkpoint(
    output_path: Path,
    ambiguous_path: Path,
    results: list[Result],
    order: dict[str, int],
):
    ordered = sorted(results, key=lambda x: order.get(x.domain, 10**9))
    write_csv(output_path, ordered)
    write_ambiguous(ambiguous_path, ordered)


def _to_int(value: str | None) -> int:
    try:
        return int(float(value or 0))
    except (TypeError, ValueError):
        return 0


def _to_float(value: str | None) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def load_existing_results(path: Path) -> list[Result]:
    """Load a previous V1/V2/V3 output for --resume; unknown/missing fields use defaults."""
    if not path.exists():
        return []

    int_fields = {
        "homepage_pet_hits", "homepage_ecom_hits", "homepage_noncommerce_hits", "homepage_blocked",
        "robots_sitemaps", "sitemap_found", "sitemap_files_scanned", "sitemap_urls_scanned",
        "sitemap_http_errors", "sitemap_parse_errors", "sitemap_gzip_files",
        "pet_urls", "commerce_urls", "pet_commerce_urls", "domain_budget_exceeded",
    }
    float_fields = {"pet_url_ratio", "pet_commerce_ratio", "score_pet", "score_ecom", "seconds"}
    valid_fields = set(asdict(Result(domain="")).keys())
    out: list[Result] = []

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            domain = normalize_domain(row.get("domain", ""))
            if not domain:
                continue
            r = Result(domain=domain)
            for key, value in row.items():
                if key not in valid_fields or key == "domain":
                    continue
                if key in int_fields:
                    setattr(r, key, _to_int(value))
                elif key in float_fields:
                    setattr(r, key, _to_float(value))
                else:
                    setattr(r, key, value or "")
            out.append(r)
    return out


def count_classes(results: list[Result]) -> dict[str, int]:
    counts = {
        "PET_FIRST_AUTO": 0,
        "PET_ADJACENT_AUTO": 0,
        "AMBIGUOUS_REVIEW": 0,
        "REJECT_AUTO": 0,
    }
    for x in results:
        counts[x.classification] = counts.get(x.classification, 0) + 1
    return counts


def format_duration(seconds: float | None) -> str:
    if seconds is None or seconds < 0 or seconds == float("inf"):
        return "--:--:--"
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def progress_text(
    completed: int,
    total: int,
    run_completed: int,
    started: float,
    results: list[Result],
) -> str:
    elapsed = max(time.monotonic() - started, 0.001)
    rate = run_completed / elapsed if run_completed > 0 else 0.0
    remaining = max(total - completed, 0)
    eta = remaining / rate if rate > 0 else None
    finish = (datetime.now() + timedelta(seconds=eta)).strftime("%H:%M:%S") if eta is not None else "--:--:--"

    width = 28
    ratio = completed / total if total else 1.0
    filled = min(width, max(0, int(width * ratio)))
    bar = "#" * filled + "-" * (width - filled)
    counts = count_classes(results)

    return (
        f"[{bar}] {ratio * 100:6.2f}%  {completed}/{total} | "
        f"{rate:5.2f} dom/s | elapsed {format_duration(elapsed)} | "
        f"ETA {format_duration(eta)} | finish ~{finish} | "
        f"P1 {counts.get('PET_FIRST_AUTO', 0)}  "
        f"Adj {counts.get('PET_ADJACENT_AUTO', 0)}  "
        f"Amb {counts.get('AMBIGUOUS_REVIEW', 0)}  "
        f"Rej {counts.get('REJECT_AUTO', 0)}"
    )


def main():
    ap = argparse.ArgumentParser(description="Low-cost Pet-first / Pet-adjacent lead classifier V3")
    ap.add_argument(
        "input",
        nargs="?",
        default=Path("data.txt"),
        type=Path,
        help="data .md/.txt/.csv containing domains (default: data.txt in same folder)",
    )
    ap.add_argument("-o", "--output", type=Path, default=Path("pet_leads_stage1.csv"))
    ap.add_argument("--workers", type=int, default=20, help="Concurrent domains (default: 20)")
    ap.add_argument("--timeout", type=float, default=8.0, help="Per-request connect/read timeout seconds (default: 8)")
    ap.add_argument("--domain-budget", type=float, default=60.0, help="Wall-clock budget per domain in seconds; 0 disables (default: 60)")
    ap.add_argument("--max-urls", type=int, default=12000, help="Max sitemap URLs sampled/domain")
    ap.add_argument("--max-sitemaps", type=int, default=36, help="Max sitemap files/domain (default: 36)")
    ap.add_argument("--checkpoint-every", type=int, default=25, help="Rewrite output CSV every N newly completed domains (default: 25)")
    ap.add_argument("--resume", action="store_true", help="Load existing output CSV and skip domains already present")
    args = ap.parse_args()

    domains = read_domains(args.input)
    if not domains:
        print("No domains found", file=sys.stderr)
        sys.exit(2)

    order = {d: i for i, d in enumerate(domains)}
    ambiguous_path = args.output.with_name(args.output.stem + "_ambiguous.txt")

    results: list[Result] = []
    completed_domains: set[str] = set()
    if args.resume and args.output.exists():
        existing = load_existing_results(args.output)
        # Keep only domains still present in the current input and only one row/domain.
        by_domain = {r.domain: r for r in existing if r.domain in order}
        results = list(by_domain.values())
        completed_domains = set(by_domain)
        print(f"Resume: loaded {len(results)} completed domains from {args.output}")

    pending_domains = [d for d in domains if d not in completed_domains]
    total = len(domains)
    resumed_count = len(results)
    print(f"Loaded {total} unique domains | already done {resumed_count} | pending {len(pending_domains)}")
    print(f"Checkpoint every {max(1, args.checkpoint_every)} domains -> {args.output}")
    print(f"Per-domain wall-clock budget: {args.domain_budget:.0f}s" if args.domain_budget > 0 else "Per-domain wall-clock budget: disabled")

    # If resume already covers everything, still normalize/checkpoint output and finish.
    if not pending_domains:
        checkpoint(args.output, ambiguous_path, results, order)
        print("Already complete.")
        print("Output:", args.output)
        print("Ambiguous:", ambiguous_path)
        return

    started = time.monotonic()
    run_completed = 0
    checkpoint_every = max(1, args.checkpoint_every)
    executor = ThreadPoolExecutor(max_workers=max(1, args.workers))
    futs = {
        executor.submit(
            analyze_domain,
            d,
            args.timeout,
            args.max_urls,
            args.max_sitemaps,
            args.domain_budget,
        ): d
        for d in pending_domains
    }

    interrupted = False
    try:
        for fut in as_completed(futs):
            d = futs[fut]
            try:
                res = fut.result()
            except Exception as e:
                res = Result(domain=d, error=f"worker: {type(e).__name__}: {str(e)[:200]}")
                res = classify(res)

            results.append(res)
            run_completed += 1
            completed = resumed_count + run_completed

            # One-line live progress with ETA. It updates after every completed domain.
            print("\r" + progress_text(completed, total, run_completed, started, results), end="", flush=True)

            # Persist the real requested output file throughout the run, not only at the end.
            if run_completed % checkpoint_every == 0 or completed == total:
                print()  # finish the progress line before checkpoint message
                checkpoint(args.output, ambiguous_path, results, order)
                print(
                    f"[checkpoint] saved {completed}/{total} rows -> {args.output} "
                    f"({datetime.now().strftime('%H:%M:%S')})",
                    flush=True,
                )

    except KeyboardInterrupt:
        interrupted = True
        print("\nCtrl+C received. Writing a checkpoint of all completed domains...", flush=True)
        checkpoint(args.output, ambiguous_path, results, order)
        print(f"Checkpoint saved: {len(results)}/{total} rows -> {args.output}", flush=True)
        for fut in futs:
            fut.cancel()
    finally:
        # On normal completion wait for workers. After Ctrl+C, do not block here waiting for slow workers.
        executor.shutdown(wait=not interrupted, cancel_futures=interrupted)

    if interrupted:
        print("You can resume later with the same command plus --resume.")
        return

    # One final normalized checkpoint, even if checkpoint-every changed or output was temporarily unavailable.
    checkpoint(args.output, ambiguous_path, results, order)
    counts = count_classes(results)
    elapsed = time.monotonic() - started

    print("\nDone")
    print("Output:", args.output)
    print("Ambiguous:", ambiguous_path)
    print("Rows:", len(results))
    print("Elapsed:", format_duration(elapsed))
    print("Counts:", counts)


if __name__ == "__main__":
    main()
