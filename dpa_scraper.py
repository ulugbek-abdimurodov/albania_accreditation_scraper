#!/usr/bin/env python3
"""DPA Albania accredited-body scraper with resumable document discovery."""
from __future__ import annotations
import argparse, json, os, re, sys, time
from pathlib import Path
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, unquote
import truststore
truststore.inject_into_ssl()
import pandas as pd
import requests
from bs4 import BeautifulSoup

SOURCE_URL = "https://dpa.gov.al/organe-te-akredituara/"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; DPAAccreditationResearch/2.0)"}
TIMEOUT = 30
RETRIES = 4
BACKOFF = 1.5
DELAY = 0.7
MEDIA_EXTS = (".pdf", ".doc", ".docx")
BASE_COLUMNS = ["company_name", "accreditation_code", "sector_scope", "address", "accreditation_date", "expiry_date", "withdrawal_date", "suspension_date", "listed_status", "scope_page_url", "source_page"]
AUDIT_COLUMNS = ["validation_status", "certificate_pdf_url", "download_path", "download_result", "document_source", "document_attempts", "discovered_document_urls", "failure_reason", "validation_note"]

def clean(v): return re.sub(r"\s+", " ", str(v or "")).strip()
def norm(v):
    x=clean(v).casefold().replace("ë","e").replace("ç","c")
    return re.sub(r"[^a-z0-9]+", " ", x).strip()
def safe_filename(v, limit=100):
    x=unquote(str(v or "")); x=re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", x); x=re.sub(r"\s+", "_", x).strip("._ ")
    return (x or "document")[:limit]
def secure_dpa_url(url):
    """Avoid slow/unreachable cleartext HTTP URLs published in older DPA rows."""
    parts=urlsplit(url)
    if parts.scheme.lower()=="http" and (parts.hostname or "").lower() in {"dpa.gov.al","www.dpa.gov.al"}:
        return urlunsplit(("https",parts.netloc,parts.path,parts.query,parts.fragment))
    return url
def build_session():
    s=requests.Session(); s.headers.update(HEADERS)
    # Use the native OS trust store (Windows Certificate Store on Windows), while
    # allowing an administrator-provided CA bundle. Never disable verification.
    s.verify=os.environ.get("REQUESTS_CA_BUNDLE") or True
    return s
def get(session, url, **kwargs):
    url=secure_dpa_url(url)
    request_timeout=kwargs.pop("timeout",TIMEOUT)
    last=None
    for attempt in range(RETRIES):
        try:
            r=session.get(url, timeout=request_timeout, **kwargs)
            if r.status_code in (408,425,429,500,502,503,504):
                last=requests.HTTPError(f"HTTP {r.status_code}", response=r)
                if attempt < RETRIES-1:
                    time.sleep(BACKOFF*(2**attempt)); continue
            r.raise_for_status(); return r
        except requests.exceptions.SSLError: raise
        except requests.HTTPError as e:
            status=e.response.status_code if e.response is not None else 0
            # A missing page/document is not transient; retrying 404s only wastes time.
            if status and status not in (408,425,429,500,502,503,504) and status < 500:
                raise
            last=e
            if attempt < RETRIES-1: time.sleep(BACKOFF*(2**attempt))
        except requests.RequestException as e:
            last=e
            if attempt < RETRIES-1: time.sleep(BACKOFF*(2**attempt))
    raise last

def find_col(headers, candidates):
    hs=[norm(h) for h in headers]
    return next((i for i,h in enumerate(hs) if any(norm(c) in h for c in candidates)), None)
def extract_tables(soup):
    out=[]
    for table in soup.find_all("table"):
        rows=table.find_all("tr"); hi=None; idx={}
        for ri,tr in enumerate(rows[:6]):
            cells=[clean(c.get_text(" ",strip=True)) for c in tr.find_all(["th","td"],recursive=False)]
            name=find_col(cells,["emri","name"]); code=find_col(cells,["kodi","code"]); scope=find_col(cells,["fusha","scope"])
            if name is not None and code is not None and scope is not None:
                hi=ri; idx={"company_name":name,"accreditation_code":code,"sector_scope":scope,
                    "address":find_col(cells,["adresa","address"]),"accreditation_date":find_col(cells,["dt e akreditimit","data e dhenies","accreditation date"]),
                    "expiry_date":find_col(cells,["data e perfundimit","expiry","expiration"]),"withdrawal_date":find_col(cells,["data e terheqjes","withdrawal"]),"suspension_date":find_col(cells,["data e pezullimit","suspension"])}; break
        if hi is None: continue
        if idx["withdrawal_date"] is not None: status="withdrawn"
        elif idx["suspension_date"] is not None: status="suspended"
        elif idx["expiry_date"] is not None: status="listed_as_accredited"
        else: status="status_unclear"
        for tr in rows[hi+1:]:
            cells=tr.find_all(["td","th"],recursive=False); vals=[clean(c.get_text(" ",strip=True)) for c in cells]
            if not vals or len(set(v.casefold() for v in vals if v))<=1: continue
            def val(k):
                i=idx.get(k); return vals[i] if i is not None and i<len(vals) else ""
            name,code=val("company_name"),val("accreditation_code")
            if not name or not code or norm(name) in {"emri","emri i ovk"}: continue
            sc=idx["sector_scope"]; a=cells[sc].find("a",href=True) if sc<len(cells) else None
            out.append({**{k:val(k) for k in BASE_COLUMNS[:8]},"listed_status":status,
                        "scope_page_url":urljoin(SOURCE_URL,a["href"]) if a else "","source_page":SOURCE_URL})
    return out

def workbook_records(path):
    if not path or not Path(path).exists(): return []
    x=pd.ExcelFile(path); sheet="All records" if "All records" in x.sheet_names else x.sheet_names[0]
    df=pd.read_excel(path,sheet_name=sheet,dtype=str).fillna("")
    return [{k:clean(row.get(k,"")) for k in BASE_COLUMNS+AUDIT_COLUMNS} for row in df.to_dict("records")]
def merge_records(fresh, saved):
    bykey={(norm(r.get("accreditation_code")),norm(r.get("scope_page_url"))):r for r in saved}
    if not fresh: return saved
    merged=[]
    for r in fresh:
        old=bykey.get((norm(r.get("accreditation_code")),norm(r.get("scope_page_url"))),{})
        merged.append({**old,**r})
        for k in AUDIT_COLUMNS:
            if k in old: merged[-1][k]=old[k]
    return merged

def media_candidates(text, base_url):
    soup=BeautifulSoup(text,"html.parser"); found=[]
    for tag in soup.find_all(True):
        for attr in ("href","src","data","data-src","data-lazy-src","data-url","content"):
            raw=tag.get(attr)
            if raw and any(ext in raw.lower() for ext in MEDIA_EXTS):
                u=urljoin(base_url,raw.strip())
                if u.startswith(("http://","https://")): found.append((u,clean(tag.get_text(" ",strip=True)),str(tag)[:500]))
    # WordPress JSON embeds and script-generated URLs, including query-bearing links.
    patterns=[r"https?://[^\s\"'<>\\]+?\.(?:pdf|docx?)(?:\?[^\s\"'<>\\]*)?",r"(?:/wp-content/uploads/)[^\s\"'<>\\]+?\.(?:pdf|docx?)(?:\?[^\s\"'<>\\]*)?"]
    for pat in patterns:
        for raw in re.findall(pat,text,re.I): found.append((urljoin(base_url,raw.replace("\\/","/")),"embedded/script URL","HTML/script"))
    # Links may have a non-file URL but disclose document intent in labels or data attributes.
    for a in soup.find_all(["a","iframe","embed","object"],href=True):
        raw=a.get("href") or a.get("src") or a.get("data")
        label=clean(a.get_text(" ",strip=True)+" "+str(a))
        if raw and any(t in label.casefold() for t in ("certifikat","certificate","akredit","scope","fusha")):
            u=urljoin(base_url,raw)
            if urlparse(u).scheme in ("http","https"): found.append((u,label,"document-labeled link"))
    unique=[]; seen=set()
    for item in found:
        if item[0] not in seen: seen.add(item[0]); unique.append(item)
    return unique

def wordpress_candidates(session, page_url, record):
    parsed=urlparse(page_url); root=f"{parsed.scheme}://{parsed.netloc}"; slug=parsed.path.strip("/").split("/")[-1]
    urls=[]
    # Public WordPress page API can expose content absent from the rendered page.
    for endpoint in (f"{root}/wp-json/wp/v2/pages?slug={slug}&_embed=1",f"{root}/wp-json/wp/v2/posts?slug={slug}&_embed=1"):
        try:
            r=get(session,endpoint,timeout=10); data=r.json()
            for item in data if isinstance(data,list) else []:
                for val in (item.get("content",{}).get("rendered",""),item.get("excerpt",{}).get("rendered",""),json.dumps(item,ensure_ascii=False)):
                    urls.extend(media_candidates(val,page_url))
        except (requests.RequestException,ValueError): pass
    # Media search is attempted only through the publicly exposed API; inaccessible APIs are ignored.
    for term in (record.get("accreditation_code",""),record.get("company_name","")):
        if not term: continue
        try:
            r=get(session,f"{root}/wp-json/wp/v2/media",params={"search":term,"per_page":30},timeout=10)
            items=r.json()
            for item in items if isinstance(items,list) else []:
                u=item.get("source_url","")
                if u: urls.append((u,item.get("title",{}).get("rendered",""),"WordPress public media API"))
        except (requests.RequestException,ValueError): pass
    return urls

def rank_candidate(item, record):
    u,label,context=item; text=norm(u+" "+label+" "+context); score=0
    code=norm(record.get("accreditation_code","")); company=norm(record.get("company_name",""))
    if code and code in text: score+=8
    if company and any(t and t in text for t in company.split() if len(t)>3): score+=2
    if any(k in text for k in ("certificate","certifikat","akredit","scope","fusha","total")): score+=3
    if any(x in u.lower() for x in (".pdf",".doc",".docx")): score+=5
    if any(k in text for k in ("suspended","withdrawn","revoked")): score-=4
    return score

def discover_document(session, record, retries=3):
    page=record.get("scope_page_url",""); attempts=0; candidates=[]; errors=[]
    if not page: return "", "", [], 0, "No detail-page URL is present"
    for attempt in range(retries):
        attempts+=1
        try:
            if any(urlparse(page).path.lower().endswith(ext) for ext in MEDIA_EXTS):
                candidates.append((page,"direct detail link","source table"))
            else:
                r=get(session,page); final=r.url
                content_type=r.headers.get("Content-Type","").lower()
                # Some DPA scope pages redirect straight to the certificate file.
                # The old code parsed the PDF bytes as HTML and missed this entirely.
                is_document=(r.content.startswith(b"%PDF") or any(urlparse(final).path.lower().endswith(ext) for ext in MEDIA_EXTS)
                             or "application/pdf" in content_type or "wordprocessingml" in content_type or "msword" in content_type)
                if is_document:
                    candidates.append((final,"detail page redirected directly to document","HTTP redirect / final response"))
                else:
                    candidates.extend(media_candidates(r.text,final))
                    # Follow pages that embed or link to another HTML detail page, one level deep.
                    soup=BeautifulSoup(r.text,"html.parser")
                    for tag in soup.find_all(["iframe","embed","object"]):
                        href=tag.get("src") or tag.get("data")
                        if href:
                            u=urljoin(final,href)
                            if not any(urlparse(u).path.lower().endswith(ext) for ext in MEDIA_EXTS):
                                try:
                                    rr=get(session,u)
                                    if rr.content.startswith(b"%PDF") or "application/pdf" in rr.headers.get("Content-Type","").lower():
                                        candidates.append((rr.url,"embedded frame redirected to PDF","HTTP redirect / final response"))
                                    else:
                                        candidates.extend(media_candidates(rr.text,rr.url))
                                except requests.RequestException as e: errors.append(str(e))
        except requests.exceptions.SSLError: raise
        except requests.RequestException as e: errors.append(str(e))
        if candidates: break
        if attempt<retries-1: time.sleep(BACKOFF*(2**attempt))
    # Query the public WordPress API once after ordinary page retries, not once
    # per retry; this avoids repeated API requests for pages without documents.
    if not candidates:
        candidates.extend(wordpress_candidates(session,page,record))
    # Prefer direct media URLs and the link with the best matching evidence.
    candidates=list({c[0]:c for c in candidates}.values())
    candidates.sort(key=lambda c:(rank_candidate(c,record), any(urlparse(c[0]).path.lower().endswith(x) for x in MEDIA_EXTS)),reverse=True)
    selected=candidates[0] if candidates else ("","","")
    why="" if candidates else (errors[-1] if errors else "No PDF or Word document link found after page and public WordPress API retries")
    return selected[0],selected[2],candidates,attempts,why

def download_document(session,url,target):
    if not url: return "","No document URL"
    try:
        with get(session,url,stream=True) as r:
            final=r.url; content_type=r.headers.get("Content-Type","").lower(); first=next(r.iter_content(8192),b"")
            if not (first.startswith(b"%PDF") or b"wordprocessingml" in content_type or b"msword" in content_type or urlparse(final).path.lower().endswith(MEDIA_EXTS)):
                return "",f"Rejected non-document response ({content_type or 'unknown content type'})"
            ext=".pdf" if first.startswith(b"%PDF") or "pdf" in content_type else (Path(urlparse(final).path).suffix.lower() or ".doc")
            target=target.with_suffix(ext); target.parent.mkdir(parents=True,exist_ok=True)
            temp=target.with_suffix(target.suffix+".part")
            with temp.open("wb") as f:
                f.write(first)
                for chunk in r.iter_content(8192):
                    if chunk: f.write(chunk)
            if temp.stat().st_size<100: temp.unlink(missing_ok=True); return "","Document response was unexpectedly small"
            temp.replace(target)
            return str(target.resolve()),"Downloaded"
    except requests.exceptions.SSLError: raise
    except requests.RequestException as e: return "",f"Download failed after retries: {e}"

def is_valid_existing(path):
    p=Path(path) if path else None
    if not p or not p.is_file() or p.stat().st_size<100: return False
    try:
        with p.open("rb") as f: head=f.read(8)
        return head.startswith(b"%PDF") or p.suffix.lower() in (".doc",".docx")
    except OSError: return False

def status_for(r,today):
    s=r.get("listed_status","")
    if s=="withdrawn": return "Withdrawn"
    if s=="suspended": return "Suspended"
    exp=pd.to_datetime(r.get("expiry_date",""),dayfirst=True,errors="coerce")
    if pd.notna(exp): return "Expired (listed as accredited; verify with DPA)" if exp.normalize()<today else "Listed as accredited; expiry date not passed"
    return "Needs manual verification"
def export(records,out_dir):
    df=pd.DataFrame(records)
    for col in BASE_COLUMNS+AUDIT_COLUMNS:
        if col not in df: df[col]=""
    df=df[BASE_COLUMNS+AUDIT_COLUMNS]
    csv=out_dir/"dpa_accredited_bodies.csv"; xlsx=out_dir/"dpa_accredited_bodies.xlsx"
    df.to_csv(csv,index=False,encoding="utf-8-sig")
    def write_workbook(path):
        with pd.ExcelWriter(path,engine="openpyxl") as writer:
            df.to_excel(writer,sheet_name="All records",index=False)
            for status,sheet in [("listed_as_accredited","Listed accredited"),("suspended","Suspended"),("withdrawn","Withdrawn"),("status_unclear","Review")]:
                sub=df[df.listed_status==status]
                if len(sub): sub.to_excel(writer,sheet_name=sheet,index=False)
            summary=pd.DataFrame([{"metric":"Total records","count":len(df)},*({"metric":k,"count":int(v)} for k,v in df.download_result.value_counts(dropna=False).items())])
            extra=out_dir/"additional_documents.csv"
            if extra.exists():
                additional=pd.read_csv(extra,dtype=str).fillna("")
                summary.loc[len(summary)]={"metric":"Additional historical documents","count":len(additional)}
                additional.to_excel(writer,sheet_name="Additional documents",index=False)
            summary.to_excel(writer,sheet_name="Download summary",index=False)
    try:
        write_workbook(xlsx)
    except PermissionError:
        # Excel often locks an open workbook. Preserve the old file and write a fresh copy.
        xlsx=out_dir/"dpa_accredited_bodies_updated.xlsx"
        write_workbook(xlsx)
    return csv,xlsx

def main():
    p=argparse.ArgumentParser(description="Extract DPA Albania accreditation records and retry certificate downloads.")
    p.add_argument("--out",default="dpa_output",help="Output folder (default: ./dpa_output)")
    p.add_argument("--input",default="dpa_accredited_bodies_input.xlsx",help="Prior workbook to preserve metadata and retry failures")
    p.add_argument("--no-download",action="store_true",help="Export metadata only")
    p.add_argument("--page-retries",type=int,default=2,help="Attempts to revisit a linked detail page (default 2)")
    args=p.parse_args(); out=Path(args.out).resolve(); out.mkdir(parents=True,exist_ok=True)
    input_path=Path(args.input)
    if args.input=="dpa_accredited_bodies_input.xlsx":
        # Resume from the newest generated report when one exists; otherwise use the supplied starter workbook.
        candidates=[out/"dpa_accredited_bodies.xlsx",out/"dpa_accredited_bodies_updated.xlsx"]
        candidates=[f for f in candidates if f.exists()]
        if candidates:
            input_path=max(candidates,key=lambda f:f.stat().st_mtime)
    saved=workbook_records(input_path); session=build_session(); records=[]
    print(f"Fetching {SOURCE_URL}")
    try:
        records=extract_tables(BeautifulSoup(get(session,SOURCE_URL).text,"html.parser"))
    except requests.exceptions.SSLError as e:
        if not saved: raise SystemExit(f"TLS verification failed; verification remains enabled. Update certifi or configure REQUESTS_CA_BUNDLE with your trusted network CA. Details: {e}")
        print(f"[WARN] DPA source page unavailable due to TLS/network issue; using saved workbook records. Details: {e}")
    except requests.RequestException as e:
        if not saved: raise SystemExit(f"Could not fetch DPA source page: {e}")
        print(f"[WARN] Source page unavailable; using saved workbook records. Details: {e}")
    records=merge_records(records,saved)
    if not records: raise SystemExit("No source records found. Check the input workbook or website layout.")
    # De-duplicate while retaining separate scoped entries for a code.
    unique={}
    for r in records: unique[(norm(r.get("accreditation_code")),norm(r.get("scope_page_url")) or norm(r.get("sector_scope")))]=r
    records=list(unique.values()); today=pd.Timestamp.today().normalize(); certdir=out/"certificates"
    for i,r in enumerate(records,1):
        r["validation_status"]=status_for(r,today)
        oldpath=r.get("download_path","")
        if args.no_download:
            r["download_result"]=r.get("download_result") or "Not requested"
        elif is_valid_existing(oldpath):
            r["download_result"]="Already present; skipped"; r["failure_reason"]=""
            if not r.get("document_source"):
                r["document_source"]="Previously downloaded local PDF; original link was not retained"
        else:
            try:
                stem=f"{safe_filename(r.get('accreditation_code'))}_{safe_filename(r.get('company_name'))}"
                expected=certdir/(stem+".pdf")
                if is_valid_existing(expected):
                    r["download_path"]=str(expected.resolve()); r["download_result"]="Already present; skipped"; r["failure_reason"]=""
                    if not r.get("document_source"):
                        r["document_source"]="Previously downloaded local PDF; original link was not retained"
                    print(f"[{i}/{len(records)}] {r.get('accreditation_code')} — {r.get('company_name')} — {r.get('download_result')}")
                    continue
                url,source,candidates,attempts,why=discover_document(session,r,max(1,args.page_retries))
                r["document_attempts"]=str(attempts); r["discovered_document_urls"]=" | ".join(c[0] for c in candidates)
                r["certificate_pdf_url"]=url; r["document_source"]=source; r["failure_reason"]=why
                if url:
                    path,result=download_document(session,url,certdir/(stem+".pdf"))
                    r["download_path"]=path; r["download_result"]=result
                    if result!="Downloaded": r["failure_reason"]=result
                else:
                    r["download_path"]=""; r["download_result"]="No document found"
                time.sleep(DELAY)
            except requests.exceptions.SSLError as e:
                r["download_result"]="TLS verification failed"; r["failure_reason"]=str(e); r["document_attempts"]=r.get("document_attempts") or "1"
            except Exception as e:
                r["download_result"]="Failed"; r["failure_reason"]=f"{type(e).__name__}: {e}"
        print(f"[{i}/{len(records)}] {r.get('accreditation_code')} — {r.get('company_name')} — {r.get('download_result','')}")
    for r in records: r["validation_note"]="Automated list/date check only; confirm official accreditation status with DPA for legal or commercial decisions."
    csv,xlsx=export(records,out)
    if args.no_download:
        print(f"\nMetadata export complete: {len(records)} records; document downloads were not attempted.")
    else:
        print(f"\nDone: {len(records)} records; {sum(r.get('download_result')=='Downloaded' for r in records)} downloaded; {sum(r.get('download_result')=='Already present; skipped' for r in records)} valid existing files skipped.")
    print(f"Excel: {xlsx}\nCSV: {csv}\nPDFs/documents: {certdir}")
if __name__=="__main__": main()

