import argparse
import csv
import logging
import re
import time
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup


USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0 Safari/537.36"


def fetch_html(url, session=None, pause=0.5, max_retries=3):
    if session is None:
        session = requests.Session()
    for attempt in range(max_retries):
        try:
            r = session.get(url, headers={"User-Agent": USER_AGENT}, timeout=20)
            r.raise_for_status()
            return r.text
        except requests.RequestException as e:
            logging.warning("fetch failed %s (attempt %s/%s): %s", url, attempt + 1, max_retries, e)
            time.sleep(pause * (attempt + 1))
    raise RuntimeError(f"Could not fetch {url} after {max_retries} attempts")


def canonicalize_link(href, base_url):
    if not href:
        return None
    if href.startswith("javascript:") or href.startswith("#"):
        return None
    return urljoin(base_url, href)


def is_cpt_code_text(text):
    if not text:
        return False
    text = text.strip()
    return bool(re.fullmatch(r"[0-9]{5}", text))


def find_code_links(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    code_links = []

    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(strip=True)
        u = canonicalize_link(href, base_url)
        if not u:
            continue

        # Prefer explicit deep code URLs in href
        m2 = re.search(r"/codes/cpt-codes/([0-9]{5})(?:/|$)", u)
        if m2:
            code = m2.group(1)
            code_links.append((code, code_page_url(code)))
            continue

        # exact code text in anchor; convert any range URL to deep code URL
        if is_cpt_code_text(text):
            code = text.strip()
            code_links.append((code, code_page_url(code)))
            continue

        # code numeric in text (e.g. label "90759")
        m = re.search(r"\b([0-9]{5})\b", text)
        if m:
            code = m.group(1)
            code_links.append((code, code_page_url(code)))
            continue

        # code from range URL pattern: /codes/cpt-codes-range/90785-90899/
        m_range = re.search(r"/codes/cpt-codes-range/([0-9]{5})(?:-[0-9]{5})?(?:/|$)", u)
        if m_range:
            code = m_range.group(1)
            code_links.append((code, code_page_url(code)))
            continue

    unique = []
    seen = set()
    for code, link in code_links:
        if code in seen:
            continue
        seen.add(code)
        unique.append((code, link))
    return unique


def find_range_links(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "/codes/cpt-codes-range/" in href:
            u = canonicalize_link(href, base_url)
            if u:
                links.append(u)
    return list(dict.fromkeys(links))


def code_page_url(code):
    return f"https://www.aapc.com/codes/cpt-codes/{code}/"


def extract_description_from_code_page(html, code):
    soup = BeautifulSoup(html, "html.parser")

    # 1) Look for direct heading text
    header = soup.find(lambda tag: tag.name in ["h1", "h2", "h3"] and code in tag.get_text())
    if header:
        # first sibling paragraphs after header
        for sib in header.next_siblings:
            if getattr(sib, "name", None) in ["p", "div"]:
                text = sib.get_text(" ", strip=True)
                if len(text) > 20:
                    return text

    # 2) standard text blocks
    for candidate in [
        soup.find("div", class_=re.compile(r"description|desc|long-descrip|cpt-description", re.I)),
        soup.find("section", class_=re.compile(r"description|desc|long-descrip", re.I)),
        soup.find("p"),
    ]:
        if candidate is not None:
            text = candidate.get_text(" ", strip=True)
            if len(text) > 20:
                return text

    # 3) meta description
    meta = soup.find("meta", attrs={"name": "description"})
    if meta and meta.get("content"):
        return meta["content"].strip()

    # 4) fallback
    cleaned = " ".join(soup.stripped_strings)
    return cleaned[:800].strip()


def crawl_range_urls(range_urls, output_csv=None, max_codes=None):
    session = requests.Session()
    results = []
    seen_codes = set()
    visited_range_urls = set()

    def extract_codes_and_ranges(html, base_url):
        codes = set()
        ranges = set()

        soup = BeautifulSoup(html, "html.parser")
        # find explicit code links
        for a in soup.find_all("a", href=True):
            href = a["href"]
            url = canonicalize_link(href, base_url)
            if not url:
                continue

            m_code = re.search(r"/codes/cpt-codes/([0-9]{5})(?:/|$)", url)
            if m_code:
                codes.add(m_code.group(1))
                continue

            m_range = re.search(r"/codes/cpt-codes-range/([0-9]{5})(?:-[0-9]{5})?(?:/|$)", url)
            if m_range:
                ranges.add(url)
                continue

                # parse text for explicit code text in anchor
            text = a.get_text(" ", strip=True)
            if is_cpt_code_text(text):
                codes.add(text)

        # no full-page numeric scraping; only parse anchor links/text to avoid false positives

        return codes, ranges

    queue = list(range_urls)
    discovered_codes = []
    discovered_code_set = set()

    while queue:
        url = queue.pop(0)
        if url in visited_range_urls:
            continue
        visited_range_urls.add(url)
        logging.info("Crawling range page: %s", url)
        try:
            html = fetch_html(url, session=session)
        except RuntimeError as e:
            logging.warning("Skipping range URL %s because fetch failed: %s", url, e)
            continue

        codes, ranges = extract_codes_and_ranges(html, url)

        # add discovered ranges to queue
        for r in ranges:
            if r not in visited_range_urls:
                queue.append(r)

        for code in codes:
            if code not in discovered_code_set:
                discovered_code_set.add(code)
                discovered_codes.append((code, code_page_url(code)))

    logging.info("Discovered %s unique codes from %s", len(discovered_codes), len(range_urls))

    if max_codes is not None:
        discovered_codes = discovered_codes[:max_codes]

    for code, code_url in discovered_codes:
        if code in seen_codes:
            continue
        logging.info("Fetching individual code page: %s -> %s", code, code_url)
        try:
            code_html = fetch_html(code_url, session=session)
        except RuntimeError as e:
            logging.warning("Skipping code %s because page cannot be fetched: %s", code, e)
            continue
        description = extract_description_from_code_page(code_html, code)
        if not description:
            logging.warning("No description found for %s", code)
        results.append({"code": code, "url": code_url, "description": description})
        seen_codes.add(code)
        time.sleep(0.2)

    if output_csv:
        with open(output_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["code", "url", "description"])
            writer.writeheader()
            for row in results:
                writer.writerow(row)

    return results


def main():
    parser = argparse.ArgumentParser(description="Extract CPT code descriptions from AAPC range pages.")
    parser.add_argument("--range-urls", nargs="+", required=True, help="Range URLs to crawl")
    parser.add_argument("--output", default="cpt_code_descriptions.csv", help="Output CSV file")
    parser.add_argument("--max-codes", type=int, default=None, help="Max codes to fetch for testing")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    results = crawl_range_urls(args.range_urls, output_csv=args.output, max_codes=args.max_codes)
    print(f"Extracted {len(results)} rows. Saved to {args.output}")


if __name__ == "__main__":
    main()
