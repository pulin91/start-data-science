"""Download the latest periodic reports for a company from SEC EDGAR.

For a ticker, this saves into ``filings/<TICKER>/``:
  * the most recent periodic report of any kind (10-Q, 10-K, 20-F, 40-F), and
  * the most recent annual report (10-K / 20-F / 40-F), used for carry-forwards
    when the interim report does not disclose an item.

Each filing is saved as the original HTML plus a PDF rendering (via headless
Chromium), so the cell notes in the comps workbook can cite physical PDF page
numbers. A ``manifest.json`` records what was fetched.

SEC requires a descriptive User-Agent with a contact e-mail; set SEC_USER_AGENT
or pass --user-agent.

Usage:
    python financial_reports/fetch_filings.py ORCL
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import date
from pathlib import Path

import requests

TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik:010d}.json"
ARCHIVE_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{acc}/{doc}"

ANNUAL_FORMS = {"10-K", "20-F", "40-F"}
PERIODIC_FORMS = ANNUAL_FORMS | {"10-Q"}

CHROME_CANDIDATES = [
    os.environ.get("CHROME_BIN", ""),
    "/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
    "chromium",
    "chromium-browser",
    "google-chrome",
]


class Edgar:
    def __init__(self, user_agent):
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": user_agent, "Accept-Encoding": "gzip, deflate"})

    def get(self, url):
        # SEC fair-access policy: max 10 requests/second.
        time.sleep(0.15)
        resp = self.session.get(url, timeout=60)
        resp.raise_for_status()
        return resp

    def cik_for(self, ticker):
        for row in self.get(TICKERS_URL).json().values():
            if row["ticker"].upper() == ticker.upper():
                return int(row["cik_str"]), row["title"]
        raise SystemExit(f"Ticker {ticker!r} not found on EDGAR")

    def periodic_filings(self, cik):
        recent = self.get(SUBMISSIONS_URL.format(cik=cik)).json()["filings"]["recent"]
        rows = [
            {
                "form": recent["form"][i],
                "filing_date": recent["filingDate"][i],
                "report_date": recent["reportDate"][i],
                "accession": recent["accessionNumber"][i],
                "primary_document": recent["primaryDocument"][i],
            }
            for i in range(len(recent["form"]))
            if recent["form"][i] in PERIODIC_FORMS
        ]
        return sorted(rows, key=lambda r: (r["filing_date"], r["report_date"]), reverse=True)


def filing_label(company, form, report_date):
    """E.g. 'Oracle QR Aug 26' or 'Oracle AR May 26' — matches the note convention."""
    kind = "AR" if form in ANNUAL_FORMS else "QR"
    d = date.fromisoformat(report_date)
    return f"{company} {kind} {d.strftime('%b %y')}"


def find_chrome():
    for c in CHROME_CANDIDATES:
        if c and (shutil.which(c) or Path(c).is_file()):
            return c
    return None


def html_to_pdf(html_path, pdf_path, chrome):
    subprocess.run(
        [
            chrome, "--headless", "--no-sandbox", "--disable-gpu",
            "--no-pdf-header-footer", f"--print-to-pdf={pdf_path}", html_path.resolve().as_uri(),
        ],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300,
    )


def download(edgar, cik, filing, out_dir, stem, chrome):
    acc = filing["accession"].replace("-", "")
    url = ARCHIVE_URL.format(cik=cik, acc=acc, doc=filing["primary_document"])
    html = edgar.get(url).text
    # Point relative image links back at EDGAR so the PDF rendering includes them.
    base = url.rsplit("/", 1)[0] + "/"
    html = html.replace("<head>", f'<head><base href="{base}">', 1)

    html_path = out_dir / f"{stem}.htm"
    html_path.write_text(html, encoding="utf-8")
    entry = {**filing, "url": url, "html": html_path.name}

    if chrome:
        pdf_path = out_dir / f"{stem}.pdf"
        html_to_pdf(html_path, pdf_path, chrome)
        entry["pdf"] = pdf_path.name
    return entry


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("ticker")
    ap.add_argument("--name", help="Short company name for labels (default: EDGAR title)")
    ap.add_argument("--out", default="filings", help="Root output folder (default: filings)")
    ap.add_argument("--user-agent", default=os.environ.get("SEC_USER_AGENT"))
    args = ap.parse_args()

    if not args.user_agent:
        sys.exit("Set SEC_USER_AGENT (e.g. 'Jane Doe jane@example.com') or pass --user-agent")

    edgar = Edgar(args.user_agent)
    cik, title = edgar.cik_for(args.ticker)
    company = args.name or title.split()[0].title()
    filings = edgar.periodic_filings(cik)
    if not filings:
        sys.exit(f"No periodic filings found for {args.ticker}")

    latest = filings[0]
    latest_annual = next((f for f in filings if f["form"] in ANNUAL_FORMS), None)

    out_dir = Path(args.out) / args.ticker.upper()
    out_dir.mkdir(parents=True, exist_ok=True)
    chrome = find_chrome()
    if not chrome:
        print("warning: Chromium not found, saving HTML only (no PDF page numbers)", file=sys.stderr)

    manifest = {"ticker": args.ticker.upper(), "cik": cik, "company": company, "filings": []}
    to_fetch = [("latest", latest)]
    if latest_annual and latest_annual["accession"] != latest["accession"]:
        to_fetch.append(("latest_annual", latest_annual))

    for role, f in to_fetch:
        stem = f"{args.ticker.upper()}_{f['form']}_{f['report_date']}"
        print(f"Downloading {f['form']} for period {f['report_date']} (filed {f['filing_date']})")
        entry = download(edgar, cik, f, out_dir, stem, chrome)
        entry["role"] = role
        entry["label"] = filing_label(company, f["form"], f["report_date"])
        manifest["filings"].append(entry)

    (out_dir / "manifest.json").write_text(json.dumps(manifest, indent=2))
    print(f"Saved to {out_dir}/")


if __name__ == "__main__":
    main()
