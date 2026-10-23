# Albania accreditation scraper

This package contains the source workbook, 124 verified PDFs downloaded from DPA record links, plus the historical 2024 Synergy-Inspect PDF supplied during this task. The updated export includes 126 accreditation records, their status lists, and download audit fields.

## Setup and run

1. Extract this folder to a writable location.
2. Open PowerShell in the extracted `dpa_accreditation_scraper` folder.
3. Install dependencies:

   ```powershell
   py -m pip install -r requirements.txt
   ```

4. Run the scraper:

   ```powershell
   py dpa_scraper.py
   ```

The script uses the newest generated workbook as its resume source when present, otherwise it uses `dpa_accredited_bodies_input.xlsx`. It writes refreshed CSV/Excel files and documents into `dpa_output`. Valid PDFs already in the certificate folder are skipped. Failed or missing records are retried on later runs.

To choose a workbook explicitly:

```powershell
py dpa_scraper.py --input "C:\path\to\previous.xlsx" --out "C:\path\to\results"
```

Metadata-only export:

```powershell
py dpa_scraper.py --no-download
```

The default detail-page retry count is 2. Increase it if DPA pages are temporarily slow:

```powershell
py dpa_scraper.py --page-retries 5
```

## What changed

- Uses Windows' trusted certificate store through `truststore` while keeping HTTPS verification enabled. If DPA publishes an old `http://dpa.gov.al/...` URL, the scraper upgrades it to HTTPS before connecting. It honors an administrator-provided `REQUESTS_CA_BUNDLE` if needed.
- Detects when a detail page redirects directly to a PDF or Word document. The previous script treated the returned PDF bytes as HTML, which caused the 21-document ceiling.
- Searches normal links, iframe/embed/object references, HTML and script-embedded media URLs, and publicly accessible WordPress page/media endpoints when ordinary page discovery fails.
- Retries transient HTTP errors and pages that initially have no document, with exponential backoff.
- Keeps an audit trail in `document_source`, `document_attempts`, `discovered_document_urls`, and `failure_reason`.
- Reuses valid local documents and exports a download summary sheet plus status-specific lists. If Excel has the standard output workbook open, it writes `dpa_accredited_bodies_updated.xlsx` instead.

## Current run results

The live DPA page returned 126 records. The package has 124 record-linked PDFs; two withdrawn/suspended records have no detail-page URL in the listing and no separate document was found. The current Synergy-Inspect row redirects to `Synergy-Dhjetor-2025.pdf`; that current document is in the record list. The 2024 PDF URL supplied by the user was also downloaded as a historical supplement and is listed on the `Additional documents` sheet and in `additional_documents.csv`.

## Validation and limitations

Syntax checks and live checks passed for the supplied Synergy PDF URL, and for ECCAT, GENIUS Lab, and Synergy detail-page redirects. Each downloaded PDF was checked for a `%PDF-` file signature. The two withdrawn/suspended rows without a detail-page URL cannot be resolved from that DPA listing; confirm with DPA if separate historical copies are required. Automated date/list checks are not an official legal status determination.
