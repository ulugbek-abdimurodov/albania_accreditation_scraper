#!/usr/bin/env python3
"""
DPA Albania accredited-body scraper
Collects accredited-body metadata and downloads linked accreditation-scope PDFs.

Install:
    python -m pip install requests beautifulsoup4 pandas openpyxl

Run:
    python dpa_scraper.py
    python dpa_scraper.py --out ./dpa_output --no-download

Use responsibly: the script requests public pages with a delay and does not bypass access controls.
"""
from __future__ import annotations

import truststore
truststore.inject_into_ssl()

import argparse
import os
import re
import ssl
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote

import pandas as pd
import requests
import certifi
from bs4 import BeautifulSoup

SOURCE_URL = "https://dpa.gov.al/organe-te-akredituara/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; PublicAccreditationResearch/1.0; contact: data-research)"
}
TIMEOUT = 25
DELAY_SECONDS = 0.6

def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()

def norm(value: str) -> str:
    value = clean(value).casefold()
    value = value.replace("ë", "e").replace("ç", "c")
    return re.sub(r"[^a-z0-9]+", " ", value).strip()

def safe_filename(value: str, max_len: int = 120) -> str:
    value = unquote(value or "")
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", "_", value).strip("._ ")
    return (value or "document")[:max_len]

def get_soup(session: requests.Session, url: str) -> BeautifulSoup:
    response = session.get(url, timeout=TIMEOUT)
    response.raise_for_status()
    response.encoding = response.apparent_encoding or response.encoding
    return BeautifulSoup(response.text, "html.parser")

def find_column(headers: list[str], candidates: list[str]) -> int | None:
    normalized = [norm(h) for h in headers]
    for i, h in enumerate(normalized):
        if any(norm(c) in h for c in candidates):
            return i
    return None

def extract_tables(soup: BeautifulSoup) -> list[dict]:
    records = []
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if not rows:
            continue

        # Find the row that most clearly acts as this table's header.
        header_idx, headers, indexes = None, [], {}
        for ri, tr in enumerate(rows[:5]):
            cells = [clean(c.get_text(" ", strip=True)) for c in tr.find_all(["th", "td"], recursive=False)]
            if not cells:
                continue
            emri = find_column(cells, ["emri", "name"])
            kodi = find_column(cells, ["kodi", "code"])
            fusha = find_column(cells, ["fusha", "scope"])
            address = find_column(cells, ["adresa", "address"])
            expiry = find_column(cells, ["data e perfundimit", "expiry", "expiration"])
            withdrawn = find_column(cells, ["data e terheqjes", "withdrawal"])
            suspended = find_column(cells, ["data e pezullimit", "suspension"])
            if emri is not None and kodi is not None and fusha is not None:
                header_idx, headers = ri, cells
                indexes = {
                    "company_name": emri,
                    "accreditation_code": kodi,
                    "scope": fusha,
                    "address": address,
                    "accreditation_date": find_column(cells, ["dt e akreditimit", "data e dhenies", "accreditation date"]),
                    "expiry_date": expiry,
                    "withdrawal_date": withdrawn,
                    "suspension_date": suspended,
                }
                break
        if header_idx is None:
            continue

        # Classify based on the columns present, not only on page headings.
        if indexes["withdrawal_date"] is not None:
            table_status = "withdrawn"
        elif indexes["suspension_date"] is not None:
            table_status = "suspended"
        elif indexes["expiry_date"] is not None:
            table_status = "listed_as_accredited"
        else:
            table_status = "status_unclear"

        for tr in rows[header_idx + 1:]:
            cells = tr.find_all(["td", "th"], recursive=False)
            if not cells:
                continue
            values = [clean(c.get_text(" ", strip=True)) for c in cells]
            # Ignore category divider rows such as "Laboratore te Akredituar".
            if len(set(v.casefold() for v in values if v)) <= 1:
                continue
            code_i = indexes["accreditation_code"]
            name_i = indexes["company_name"]
            if max(code_i, name_i) >= len(values):
                continue
            code, name = values[code_i], values[name_i]
            if not code or not name or norm(code) == "kodi" or norm(name) in {"emri i ovk", "emri"}:
                continue

            def value_at(key):
                i = indexes.get(key)
                return values[i] if i is not None and i < len(values) else ""

            scope_i = indexes["scope"]
            scope_cell = cells[scope_i] if scope_i < len(cells) else None
            scope_link = ""
            if scope_cell:
                a = scope_cell.find("a", href=True)
                if a:
                    scope_link = urljoin(SOURCE_URL, a["href"])

            records.append({
                "company_name": name,
                "accreditation_code": code,
                "sector_scope": value_at("scope"),
                "address": value_at("address"),
                "accreditation_date": value_at("accreditation_date"),
                "expiry_date": value_at("expiry_date"),
                "withdrawal_date": value_at("withdrawal_date"),
                "suspension_date": value_at("suspension_date"),
                "listed_status": table_status,
                "scope_page_url": scope_link,
                "source_page": SOURCE_URL,
            })
    return records

def find_pdf_url(session: requests.Session, url: str) -> str:
    """Return direct PDF URL or the first PDF linked from a DPA detail page."""
    if not url:
        return ""
    if urlparse(url).path.lower().endswith(".pdf"):
        return url
    try:
        soup = get_soup(session, url)
        links = []
        for a in soup.find_all("a", href=True):
            href = urljoin(url, a["href"])
            path = urlparse(href).path.lower()
            if path.endswith(".pdf"):
                links.append(href)
        # Prefer links with terms likely to identify the current certificate/scope.
        if links:
            preferred = [u for u in links if any(k in u.casefold() for k in ("total", "akredit", "certificate", "certifikat", "scope"))]
            return (preferred or links)[0]
    except requests.RequestException as exc:
        print(f"[WARN] Could not inspect document page {url}: {exc}")
    return ""

def status_for(record: dict, today: pd.Timestamp) -> str:
    listed = record["listed_status"]
    if listed == "withdrawn":
        return "Withdrawn"
    if listed == "suspended":
        return "Suspended"
    expiry = pd.to_datetime(record.get("expiry_date", ""), dayfirst=True, errors="coerce")
    if pd.notna(expiry):
        if expiry.normalize() < today.normalize():
            return "Expired (listed as accredited; verify with DPA)"
        return "Listed as accredited; expiry date not passed"
    return "Needs manual verification"

def download_pdf(session: requests.Session, url: str, target: Path) -> tuple[str, str]:
    if not url:
        return "", "No PDF link found"
    try:
        response = session.get(url, timeout=TIMEOUT, stream=True)
        response.raise_for_status()
        content_type = response.headers.get("Content-Type", "").lower()
        first_chunk = next(response.iter_content(chunk_size=8192), b"")
        if not (first_chunk.startswith(b"%PDF") or "pdf" in content_type):
            return "", f"Skipped: response does not appear to be a PDF ({content_type or 'unknown type'})"
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("wb") as f:
            f.write(first_chunk)
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        return str(target), "Downloaded"
    except requests.RequestException as exc:
        return "", f"Download failed: {exc}"

def main():
    parser = argparse.ArgumentParser(description="Collect DPA accredited-body records and certificate PDFs.")
    parser.add_argument("--out", default="dpa_output", help="Output folder (default: dpa_output)")
    parser.add_argument("--no-download", action="store_true", help="Only export metadata; do not download PDFs")
    args = parser.parse_args()
    out_dir = Path(args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.headers.update(HEADERS)
    # Use an up-to-date trusted CA bundle. If your network uses an approved
    # HTTPS-inspecting proxy, set REQUESTS_CA_BUNDLE to your organization's CA bundle.
    session.verify = os.environ.get("REQUESTS_CA_BUNDLE") or certifi.where()

    print(f"Fetching {SOURCE_URL}")
    try:
        soup = get_soup(session, SOURCE_URL)
    except requests.exceptions.SSLError as exc:
        raise SystemExit(
            "\\nSSL certificate verification failed. Do not use verify=False.\\n"
            "Try: python -m pip install --upgrade certifi requests\\n"
            "Then retry. If you are on a corporate/school proxy or antivirus HTTPS inspection, "
            "configure REQUESTS_CA_BUNDLE to the trusted CA certificate supplied by your IT/admin.\\n"
            f"Details: {exc}"
        ) from exc
    except requests.exceptions.RequestException as exc:
        raise SystemExit(
            f"Could not fetch {SOURCE_URL}. Check internet access, DNS, proxy, and website availability.\\n"
            f"Details: {exc}"
        ) from exc
    records = extract_tables(soup)
    if not records:
        raise SystemExit("No matching tables found. The DPA page layout may have changed; inspect the page HTML.")

    # De-duplicate by code + scope link while retaining multiple accreditation schemes for one company.
    unique = {}
    for r in records:
        key = (norm(r["accreditation_code"]), r["scope_page_url"] or norm(r["sector_scope"]))
        unique[key] = r
    records = list(unique.values())

    today = pd.Timestamp.today().normalize()
    for i, r in enumerate(records, start=1):
        r["validation_status"] = status_for(r, today)
        r["certificate_pdf_url"] = ""
        r["download_path"] = ""
        r["download_result"] = "Not requested" if args.no_download else "Pending"
        if not args.no_download:
            pdf_url = find_pdf_url(session, r["scope_page_url"])
            r["certificate_pdf_url"] = pdf_url
            code = safe_filename(r["accreditation_code"])
            company = safe_filename(r["company_name"])
            if pdf_url:
                target = out_dir / "certificates" / f"{code}_{company}.pdf"
                path, result = download_pdf(session, pdf_url, target)
                r["download_path"], r["download_result"] = path, result
            else:
                r["download_result"] = "No PDF link found on linked page"
            time.sleep(DELAY_SECONDS)
        print(f"[{i}/{len(records)}] {r['accreditation_code']} — {r['company_name']} — {r['validation_status']}")

    df = pd.DataFrame(records)
    # Add a human-friendly note so users don't confuse date calculation with official verification.
    df["validation_note"] = df["validation_status"].apply(
        lambda s: "Automated date/list check only; confirm official status with DPA for legal/commercial decisions."
    )
    csv_path = out_dir / "dpa_accredited_bodies.csv"
    xlsx_path = out_dir / "dpa_accredited_bodies.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All records", index=False)
        for status, sheet in [
            ("listed_as_accredited", "Listed accredited"),
            ("suspended", "Suspended"),
            ("withdrawn", "Withdrawn"),
            ("status_unclear", "Review"),
        ]:
            subset = df[df["listed_status"] == status]
            if not subset.empty:
                subset.to_excel(writer, sheet_name=sheet, index=False)

    print(f"\nFinished. Records: {len(df)}")
    print(f"Excel: {xlsx_path}")
    print(f"CSV:   {csv_path}")
    if not args.no_download:
        print(f"PDFs:  {out_dir / 'certificates'}")

if __name__ == "__main__":
    main()
