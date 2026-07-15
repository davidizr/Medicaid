import argparse
import csv
import logging
import re
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0 Safari/537.36"


def fetch_html(url, session=None, pause=0.4, max_retries=3):
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


def code_page_url(code):
    code = str(code).strip()
    if not code:
        return None
    code = code.strip().upper()
    # Only letters and digits in slug
    slug = re.sub(r"[^A-Z0-9]", "", code)
    if not slug:
        return None
    return f"https://www.aapc.com/codes/cpt-codes/{slug}/"


def extract_description_from_code_page(html, code):
    soup = BeautifulSoup(html, "html.parser")

    # Try obvious heading + paragraph description
    header = soup.find(lambda tag: tag.name in ["h1", "h2", "h3"] and code in tag.get_text())
    if header:
        for sib in header.next_siblings:
            if getattr(sib, "name", None) in ["p", "div"]:
                text = sib.get_text(" ", strip=True)
                if len(text) > 20:
                    return text

    # Try common description containers
    for candidate in [
        soup.find("div", class_=re.compile(r"description|desc|long-descrip|cpt-description", re.I)),
        soup.find("section", class_=re.compile(r"description|desc|long-descrip", re.I)),
        soup.find("p"),
    ]:
        if candidate is not None:
            text = candidate.get_text(" ", strip=True)
            if len(text) > 20:
                return text

    meta = soup.find("meta", attrs={"name": "description"})
    if meta and meta.get("content"):
        return meta["content"].strip()

    cleaned = " ".join(soup.stripped_strings)
    return cleaned[:800].strip()


def truthy_value(val):
    if val is None:
        return False
    if isinstance(val, bool):
        return val
    val = str(val).strip().lower()
    if not val:
        return False
    if val in {"1", "true", "t", "yes", "y", "on"}:
        return True
    try:
        return float(val) != 0
    except Exception:
        return False


def detect_columns(header):
    header_norm = [h.strip().upper() for h in header]
    cardio_options = ["IS_CARDIOVASCULAR", "CARDIO", "IS_CARDIO", "IS_CV"]
    neuro_options = ["IS_NEUROLOGY", "NEURO", "IS_NEURO", "IS_CNS"]
    cardio = next((h for h in header if h.strip().upper() in cardio_options), None)
    neuro = next((h for h in header if h.strip().upper() in neuro_options), None)
    return cardio, neuro


def load_codes_from_csv(input_csv, code_column, cardio_column, neuro_column):
    codes = set()
    with open(input_csv, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if code_column not in reader.fieldnames:
            raise ValueError(f"Code column '{code_column}' not found in input CSV columns: {reader.fieldnames}")

        if cardio_column is not None and cardio_column not in reader.fieldnames:
            raise ValueError(f"Cardio column '{cardio_column}' not found in input CSV columns: {reader.fieldnames}")
        if neuro_column is not None and neuro_column not in reader.fieldnames:
            raise ValueError(f"Neuro column '{neuro_column}' not found in input CSV columns: {reader.fieldnames}")

        for row in reader:
            is_cardio = truthy_value(row.get(cardio_column)) if cardio_column else False
            is_neuro = truthy_value(row.get(neuro_column)) if neuro_column else False
            if is_cardio or is_neuro:
                code = row.get(code_column, "").strip()
                if code:
                    codes.add(code)
    return codes


def main():
    parser = argparse.ArgumentParser(description="Lookup CPT code descriptions for cardio/neuro codes from AAPC code pages.")
    parser.add_argument("--input", required=True, help="Input CSV file containing HCPCS/CPT codes and flags.")
    parser.add_argument("--output", default="lookup_code_descriptions.csv", help="Output CSV file with code,description.")
    parser.add_argument("--code-column", default="HCPCS_CODE", help="Name of code column. Default HCPCS_CODE.")
    parser.add_argument("--cardio-column", default="IS_CARDIOVASCULAR", help="Cardio binary column name.")
    parser.add_argument("--neuro-column", default="IS_NEUROLOGY", help="Neuro binary column name.")
    parser.add_argument("--max-codes", type=int, default=None, help="Optional max number of codes to process.")
    parser.add_argument("--pause", type=float, default=0.2, help="Pause seconds between requests.")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    logging.info("Reading codes from %s", args.input)
    codes = load_codes_from_csv(args.input, args.code_column, args.cardio_column, args.neuro_column)
    codes = sorted(codes)
    logging.info("Found %s unique cardio/neuro codes", len(codes))
    if args.max_codes is not None:
        codes = codes[: args.max_codes]

    session = requests.Session()
    results = []
    for code in codes:
        url = code_page_url(code)
        if not url:
            logging.warning("Skipping invalid code '%s'", code)
            continue
        logging.info("Fetching %s", url)
        try:
            html = fetch_html(url, session=session)
            desc = extract_description_from_code_page(html, code)
            if not desc:
                logging.warning("No description found for %s", code)
                desc = ""
            results.append({"code": code, "description": desc})
        except Exception as e:
            logging.warning("Failed %s: %s", code, e)
            results.append({"code": code, "description": ""})
        time.sleep(args.pause)

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["code", "description"])
        writer.writeheader()
        for row in results:
            writer.writerow(row)

    logging.info("Wrote %s rows to %s", len(results), args.output)


if __name__ == "__main__":
    main()
