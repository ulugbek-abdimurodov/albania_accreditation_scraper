# DPA Albania Accreditation Scraper

Collects public accreditation-body records from https://dpa.gov.al/organe-te-akredituara/, exports CSV and Excel, and attempts to download PDF scope/certificate documents linked from each row.

## Setup
Requires Python 3.10+.

```bash
python -m pip install --upgrade requests certifi beautifulsoup4 pandas openpyxl
```

## Run
```bash
python dpa_scraper.py
```

This creates `dpa_output/` with:
- `dpa_accredited_bodies.xlsx`
- `dpa_accredited_bodies.csv`
- `certificates/` (downloaded PDFs)

Metadata-only run:
```bash
python dpa_scraper.py --no-download
```

Custom output folder:
```bash
python dpa_scraper.py --out ./output
```

## Validation caveats
The script distinguishes records in the page's accredited, suspended, and withdrawn tables, and checks expiry dates. An expired date does not necessarily mean DPA has formally withdrawn accreditation, so those records are marked for verification. Automated extraction cannot replace confirmation from DPA.

## Notes
- The page may change its HTML structure; review the output and logs after running.
- Some `Fusha e Akreditimit` links lead to a detail page rather than a direct PDF. The script searches that page for a PDF link and downloads the first likely match. If a page links to multiple PDFs, verify that the selected PDF is the correct current scope.
- Requests are paced; no access controls are bypassed.

## SSL certificate error on Windows

If you see `SSLCertVerificationError: unable to get local issuer certificate`:

1. Run `python -m pip install --upgrade requests certifi`.
2. Retry the scraper.
3. If it still fails, check whether your company/school network or antivirus inspects HTTPS. Ask the administrator for the trusted root CA certificate and set `REQUESTS_CA_BUNDLE` to its PEM file path. Do not disable certificate verification with `verify=False`.
