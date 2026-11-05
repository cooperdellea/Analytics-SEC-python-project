"""Dynamic SEC 10-K financial and sentiment research tool.

Install: py -m pip install requests beautifulsoup4 openpyxl yfinance
Run:     py "$HOME/Downloads/sec_financial_analysis_two_charts.py"

Sources: SEC EDGAR XBRL Company Facts and 10-K HTML; Yahoo Finance daily close;
user-provided Loughran-McDonald master dictionary CSV.

Outputs are research aids. Validate XBRL tags, MD&A boundaries, and prices against
original sources before submitting financial conclusions.
"""

from __future__ import annotations

import csv
import re
import sys
import time
import unicodedata
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
import yfinance as yf
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.axis import ChartLines
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
ARCHIVES_URL = "https://www.sec.gov/Archives/edgar/data"
CATEGORIES = ("Negative", "Positive", "Uncertainty", "Litigious",
              "Strong_Modal", "Weak_Modal", "Constraining")
NAVY = "17365D"
LIGHT = "EAF1F8"


STATEMENT_TAGS = {
    "Income Statement": {
        "Revenue": ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet", "RevenueFromContractWithCustomerIncludingAssessedTax"],
        "Cost of revenue": ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsSold"],
        "Gross profit": ["GrossProfit"],
        "Research and development": ["ResearchAndDevelopmentExpense"],
        "Operating expenses": ["OperatingExpenses"],
        "Operating income": ["OperatingIncomeLoss"],
        "Income before tax": ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
        "Income tax expense": ["IncomeTaxExpenseBenefit"],
        "Net income": ["NetIncomeLoss", "ProfitLoss"],
    },
    "Balance Sheet": {
        "Cash and equivalents": ["CashAndCashEquivalentsAtCarryingValue"],
        "Short-term investments": ["ShortTermInvestments", "MarketableSecuritiesCurrent"],
        "Accounts receivable": ["AccountsReceivableNetCurrent"],
        "Inventory": ["InventoryNet"],
        "Current assets": ["AssetsCurrent"],
        "Total assets": ["Assets"],
        "Current liabilities": ["LiabilitiesCurrent"],
        "Long-term debt": ["LongTermDebtNoncurrent", "LongTermDebt"],
        "Total liabilities": ["Liabilities"],
        "Stockholders' equity": ["StockholdersEquity"],
        "Liabilities and equity": ["LiabilitiesAndStockholdersEquity"],
    },
    "Cash Flows": {
        "Net income": ["NetIncomeLoss"],
        "Depreciation and amortization": ["DepreciationDepletionAndAmortization", "DepreciationDepletionAndAmortizationPropertyPlantAndEquipment", "DepreciationAmortizationAndAccretionNet", "DepreciationAndAmortizationInCashFlowStatement", "DepreciationDepletionAndAmortizationPropertyPlantAndEquipmentIncludingAmortizationOfIntangibleAssets"],
        "Operating cash flow": ["NetCashProvidedByUsedInOperatingActivities"],
        "Capital expenditures": ["PaymentsToAcquirePropertyPlantAndEquipment"],
        "Investing cash flow": ["NetCashProvidedByUsedInInvestingActivities"],
        "Financing cash flow": ["NetCashProvidedByUsedInFinancingActivities"],
        "Change in cash": ["CashAndCashEquivalentsPeriodIncreaseDecrease", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect"],
    },
}


class SecClient:
    """Reusable SEC session with timeout, throttling, and helpful HTTP errors."""

    def __init__(self, identity):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": identity,
            "Accept-Encoding": "gzip, deflate",
            "Accept": "application/json, text/html, */*",
        })
        self.last_request = 0.0

    def get(self, url, json_response=True):
        # Keep well below the SEC's published maximum request rate.
        wait = 0.25 - (time.monotonic() - self.last_request)
        if wait > 0:
            time.sleep(wait)
        try:
            response = self.session.get(url, timeout=40)
            self.last_request = time.monotonic()
            response.raise_for_status()
            return response.json() if json_response else response.content
        except (requests.RequestException, ValueError) as error:
            raise RuntimeError(f"SEC request failed ({url}): {error}") from error

def lookup_company(client, ticker):
    """Return SEC CIK and registered name for an exact ticker match."""
    data = client.get(TICKERS_URL)
    for company in data.values():
        if company.get("ticker", "").upper() == ticker:
            return int(company["cik_str"]), company["title"]
    raise ValueError(f"Ticker {ticker} not found in SEC company ticker data.")

def recent_10ks(client, cik, count):
    """Get recent 10-K filing records, including older submissions if necessary."""
    data = client.get(SUBMISSIONS_URL.format(cik=f"{cik:010d}"))
    recent = data.get("filings", {}).get("recent", {})
    records = [dict(zip(recent.keys(), values)) for values in zip(*recent.values())]
    for older in data.get("filings", {}).get("files", []):
        if len([r for r in records if r.get("form") == "10-K"]) >= count:
            break
        more = client.get("https://data.sec.gov/submissions/" + older["name"])
        records.extend(dict(zip(more.keys(), values)) for values in zip(*more.values()))
    filings = [r for r in records if r.get("form") == "10-K"]
    filings.sort(key=lambda item: item["filingDate"], reverse=True)
    return filings[:count]

def filing_year(filing):
    """Use fiscal report year rather than submission year for labels."""
    return filing["reportDate"][:4]

def download_filing(client, cik, ticker, filing, folder):
    """Save original SEC primary HTML document with a unique accession suffix."""
    accession = filing["accessionNumber"].replace("-", "")
    primary = filing["primaryDocument"]
    url = f"{ARCHIVES_URL}/{cik}/{accession}/{primary}"
    suffix = Path(primary).suffix.lower() or ".html"
    filename = f"{ticker}_10-K_{filing_year(filing)}_{accession}{suffix}"
    destination = folder / filename
    if not destination.exists():
        destination.write_bytes(client.get(url, json_response=False))
    return destination, url

def extract_mda(html_bytes):
    """Extract Item 7 from visible HTML text using section-heading candidates.

    Returns None if section boundaries are ambiguous. This is intentionally
    conservative: SEC HTML layouts differ and table-of-contents headings may
    resemble real headings.
    """
    soup = BeautifulSoup(html_bytes, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = soup.get_text("\n", strip=True)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    # Require both the Item 7 label and its MD&A title to avoid most TOC hits.
    start_pattern = re.compile(
        r"(?im)^\s*item\s*7\s*[.\-:–]?\s*"
        r"management[’'s\s]+discussion\s+and\s+analysis",
    )
    end_pattern = re.compile(r"(?im)^\s*item\s*(?:7a|8)\s*[.\-:–]?\s*")
    starts = list(start_pattern.finditer(text))
    ends = list(end_pattern.finditer(text))
    candidates = []
    for start in starts:
        next_ends = [end for end in ends if end.start() > start.end()]
        if not next_ends:
            continue
        segment = text[start.start():next_ends[0].start()].strip()
        # Short matches are generally table-of-contents entries.
        if len(segment) >= 1500:
            candidates.append(segment)
    if not candidates:
        return None
    # Prefer the longest credible section; still flag extraction for review.
    return max(candidates, key=len)

def save_mda(html_path, ticker, filing, folder):
    """Write extracted MD&A text to UTF-8, or return None if not found."""
    extracted = extract_mda(html_path.read_bytes())
    if not extracted:
        return None
    accession = filing["accessionNumber"].replace("-", "")
    destination = folder / f"{ticker}_MDA_{filing_year(filing)}_{accession}.txt"
    destination.write_text(extracted, encoding="utf-8")
    return destination

def find_fact(facts, tags, filing, instant):
    """Find a USD fact for a particular 10-K accession and fiscal report date.

    Flow items must span roughly one fiscal year; balance sheet items are
    instant facts. Matching accession prevents accidentally mixing restatements
    from later filings. Blank means no confidently matched standardized tag.
    """
    us_gaap = facts.get("facts", {}).get("us-gaap", {})
    report_date = filing["reportDate"]
    accession = filing["accessionNumber"]
    for tag in tags:
        entries = us_gaap.get(tag, {}).get("units", {}).get("USD", [])
        matches = []
        for entry in entries:
            if entry.get("accn") != accession or entry.get("end") != report_date:
                continue
            if entry.get("form") != "10-K":
                continue
            if instant:
                if "start" in entry:
                    continue
            else:
                start = entry.get("start")
                if not start:
                    continue
                try:
                    duration = (datetime.fromisoformat(report_date) -
                                datetime.fromisoformat(start)).days
                except ValueError:
                    continue
                if not 330 <= duration <= 380:
                    continue
            matches.append(entry)
        if matches:
            # Prefer the latest frame/filing match when duplicates exist.
            matches.sort(key=lambda e: (bool(e.get("frame")), e.get("filed", "")),
                         reverse=True)
            return matches[0]["val"]
    return None

def clean_mda_text(raw_text):
    """Return lowercase, word-only text suitable for dictionary matching.

    This intentionally does NOT remove stop words or stem words, because
    doing so could change sentiment interpretation and dictionary matches.
    Numeric amounts and table figures are omitted, but table labels remain.
    """
    if not isinstance(raw_text, str):
        raise TypeError("MD&A content must be a string.")

    # Remove known page footers before merging lines.
    lines = [
        line for line in raw_text.splitlines()
        if not PAGE_FOOTER.fullmatch(line.strip())
    ]
    text = "\n".join(lines)

    # Convert typographic apostrophes to standard apostrophes, then remove
    # possessive endings so "Company's" is counted as "company".
    text = text.replace("\u2019", "'").replace("\u2018", "'")
    text = re.sub(r"(?i)\b([a-z]+)'s\b", r"\1", text)

    # Unicode normalization converts many accented Latin characters into
    # base letters. Unrecognized symbols are ignored in the word-only output.
    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        char for char in text if not unicodedata.combining(char)
    ).lower()

    # Extract words, removing numeric tokens and punctuation. Do not filter
    # stop words, since words such as "not" are important for later analysis.
    words = WORD_PATTERN.findall(text)
    words = [word for word in words if len(word) > 1 or word in {"a", "i"}]
    return " ".join(words)

def read_dictionary(csv_path):
    """Load positive category-year flags from the official master CSV."""
    words = {category: set() for category in CATEGORIES}
    with csv_path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if not reader.fieldnames:
            raise ValueError("Dictionary CSV is empty or missing headers.")
        columns = {name.strip().lower().replace("_", ""): name
                   for name in reader.fieldnames}
        required = ["Word", *CATEGORIES]
        missing = [name for name in required
                   if name.lower().replace("_", "") not in columns]
        if missing:
            raise ValueError(f"Dictionary is missing columns: {missing}")
        for row in reader:
            word = (row[columns["word"]] or "").strip().lower()
            if not word or not word.isalpha():
                continue
            for category in CATEGORIES:
                key = columns[category.lower().replace("_", "")]
                try:
                    # Positive values mean active category membership;
                    # negative values indicate removal in dictionary updates.
                    flag = int(float((row[key] or "0").strip()))
                except ValueError as exc:
                    raise ValueError(
                        f"Invalid {category} flag for {word!r}"
                    ) from exc
                if flag > 0:
                    words[category].add(word)
    if not words["Positive"] or not words["Negative"]:
        raise ValueError("Dictionary lacks positive or negative word entries.")
    return words


PAGE_FOOTER = re.compile(
    r"^\s*.+?\s*\|\s*20\d{2}\s+Form\s+10-K\s*\|\s*\d+\s*$",
    re.IGNORECASE,
)
WORD_PATTERN = re.compile(r"[a-z]+")


def prompt_settings():
    """Collect company, filing count, output location, SEC ID, and dictionary."""
    while True:
        ticker = input("Company ticker (e.g., AAPL): ").strip().upper()
        if re.fullmatch(r"[A-Z0-9]+(?:[.-][A-Z0-9]+)*", ticker):
            break
        print("Enter a valid ticker symbol.")
    while True:
        try:
            count = int(input("Number of recent 10-Ks (1-10): ").strip())
            if 1 <= count <= 10:
                break
        except ValueError:
            pass
        print("Please enter a whole number from 1 to 10.")
    while True:
        raw = input("Save folder (full path): ").strip().strip('"')
        if not raw:
            print("A save folder is required.")
            continue
        try:
            folder = Path(raw).expanduser().resolve()
            folder.mkdir(parents=True, exist_ok=True)
            break
        except OSError as error:
            print(f"Cannot use folder: {error}")
    while True:
        identity = input("Your name and email for SEC requests: ").strip()
        if re.search(r"[^\s@]+@[^\s@]+\.[^\s@]+", identity):
            break
        print("Please enter an identifying name and valid email address.")
    while True:
        raw = input("Loughran-McDonald master CSV full path: ").strip().strip('"')
        dictionary_path = Path(raw).expanduser()
        if dictionary_path.is_file() and dictionary_path.suffix.lower() == ".csv":
            break
        print("The dictionary CSV was not found. Enter its full path.")
    return ticker, count, folder, identity, dictionary_path


def get_prior_assets(facts, year, latest_filing_date):
    """Find preceding fiscal-year assets available by the current filing date.

    This is used only for the first requested year's asset turnover. Reject
    facts published after the current year's filing to avoid look-ahead bias.
    """
    entries = (facts.get("facts", {}).get("us-gaap", {})
               .get("Assets", {}).get("units", {}).get("USD", []))
    options = [item for item in entries
               if item.get("end", "").startswith(f"{year - 1}-")
               and item.get("form") == "10-K"
               and item.get("filed", "9999") <= latest_filing_date
               and "start" not in item]
    if not options:
        return None
    options.sort(key=lambda x: (x.get("filed", ""), x.get("accn", "")),
                 reverse=True)
    return options[0]["val"]


def score_text(raw_text, dictionary):
    """Clean text and count exact Loughran-McDonald category word matches."""
    if not re.match(r"^\s*Item\s+7\s*[.:]", raw_text, re.I):
        raise ValueError("MD&A text does not begin with Item 7")
    if re.search(r"(?im)^\s*Item\s+(?:7A|8)\s*[.:]", raw_text):
        raise ValueError("MD&A appears to include Item 7A or Item 8")
    cleaned = clean_mda_text(raw_text)
    tokens = Counter(cleaned.split())
    if not tokens:
        raise ValueError("MD&A contains no words after cleaning")
    counts = {cat: sum(tokens[word] for word in dictionary[cat])
              for cat in CATEGORIES}
    return cleaned, sum(tokens.values()), counts


def fetch_prices(ticker, filings):
    """Get unadjusted daily closing prices on/before each fiscal year-end.

    Yahoo data is not SEC-sourced. Record the actual trading date and verify
    the figures against a second price source for final research use.
    """
    dates = [date.fromisoformat(f["reportDate"]) for f in filings]
    start = min(dates) - timedelta(days=12)
    end = max(dates) + timedelta(days=1)
    history = yf.Ticker(ticker).history(
        start=start.isoformat(), end=end.isoformat(),
        interval="1d", auto_adjust=False, actions=False,
    )
    if history.empty or "Close" not in history.columns:
        raise RuntimeError("Yahoo Finance returned no closing prices")
    sessions = [(timestamp.date(), float(row["Close"]))
                for timestamp, row in history.iterrows()
                if row["Close"] is not None
                and float(row["Close"]) > 0
                and float(row["Close"]) == float(row["Close"])]
    result = {}
    for filing, balance_date in zip(filings, dates):
        eligible = [(day, close) for day, close in sessions
                    if day <= balance_date and (balance_date - day).days <= 10]
        if not eligible:
            print(f"WARNING: No recent closing price for {balance_date}")
            continue
        result[int(filing_year(filing))] = (balance_date,
                                           max(eligible, key=lambda x: x[0]))
    return result


def style_sheet(sheet, header_row=1, first_data_row=2, widths=None):
    """Apply consistent formatting, freeze panes, filtering, and banded rows."""
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = f"B{first_data_row}"
    sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(sheet.max_column)}{sheet.max_row}"
    sheet.row_dimensions[header_row].height = 30
    for cell in sheet[header_row]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    for row in sheet.iter_rows(min_row=first_data_row):
        for cell in row:
            cell.font = Font(name="Aptos", size=10, color="243746")
            cell.alignment = Alignment(vertical="center")
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F6FA")
    if widths:
        for idx, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(idx)].width = width
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.print_options.horizontalCentered = True


def add_financial_sheets(wb, facts, filings, company, ticker):
    """Write one year per column with XBRL values from its original 10-K."""
    filings = sorted(filings, key=lambda filing: filing["reportDate"])
    positions = {}
    for name, tags_by_label in STATEMENT_TAGS.items():
        sheet = wb.create_sheet(name)
        sheet.cell(1, 1, f"{company} ({ticker}) | {name}")
        sheet.cell(2, 1, "USD | original 10-K XBRL facts | blanks = unavailable")
        sheet.cell(4, 1, "Financial line item")
        for col, filing in enumerate(filings, 2):
            sheet.cell(4, col, int(filing_year(filing)))
            sheet.cell(5, col, filing["reportDate"])
        for row, (label, tags) in enumerate(tags_by_label.items(), 6):
            sheet.cell(row, 1, label)
            for col, filing in enumerate(filings, 2):
                value = find_fact(facts, tags, filing, name == "Balance Sheet")
                if name == "Income Statement" and label == "Operating expenses" and value is None:
                    gross = find_fact(facts, STATEMENT_TAGS[name]["Gross profit"], filing, False)
                    operating_income = find_fact(facts, STATEMENT_TAGS[name]["Operating income"], filing, False)
                    if gross is not None and operating_income is not None:
                        # Derived figure, not an independently reported XBRL fact.
                        value = gross - operating_income
                        sheet.cell(row, col).comment = Comment(
                            "DERIVED: Gross profit minus operating income; "
                            "not a directly retrieved OperatingExpenses XBRL fact.",
                            "Data provenance")
                sheet.cell(row, col, value)
                sheet.cell(row, col).number_format = '#,##0;[Red](#,##0);–'
                positions[(name, label, int(filing_year(filing)))] = (row, col, value)
        style_sheet(sheet, header_row=4, first_data_row=6,
                    widths=[37] + [19] * len(filings))
        sheet.freeze_panes = "B6"
        sheet["A1"].font = Font(name="Aptos Display", size=16,
                                color=NAVY, bold=True)
        sheet["A2"].font = Font(name="Aptos", size=10, color="667788")
        sheet.row_dimensions[1].height = 27
        sheet.auto_filter.ref = f"A4:{get_column_letter(len(filings)+1)}{sheet.max_row}"
        sheet.sheet_properties.tabColor = NAVY
    return positions


def fact_ref(positions, sheet, label, year):
    """Return a quoted Excel cell reference only when the fact is numeric."""
    item = positions.get((sheet, label, year))
    if item and isinstance(item[2], (int, float)):
        return f"'{sheet}'!{get_column_letter(item[1])}{item[0]}"
    return None


def add_ratios(wb, positions, filings, facts):
    """Build ratio formulas with missing-data safeguards and prior-year assets."""
    sheet = wb.create_sheet("Financial Ratios")
    sheet.append(["Fiscal Year", "Current Ratio", "Liabilities / Assets",
                  "Asset Turnover", "Net Profit Margin", "Data Notes"])
    years = sorted((int(filing_year(f)) for f in filings))
    first_year = years[0]
    prior_assets = get_prior_assets(facts, first_year,
                                   next(f["filingDate"] for f in filings
                                        if int(filing_year(f)) == first_year))
    # Supporting year is included only for efficiency ratio calculation.
    if prior_assets is not None:
        sheet["G2"] = f"Prior-year assets ({first_year - 1}; USD)"
        sheet["H2"] = prior_assets
        sheet["H2"].number_format = '#,##0;[Red](#,##0)'
    for year in years:
        row = sheet.max_row + 1
        sheet.cell(row, 1, year)
        ca = fact_ref(positions, "Balance Sheet", "Current assets", year)
        cl = fact_ref(positions, "Balance Sheet", "Current liabilities", year)
        liab = fact_ref(positions, "Balance Sheet", "Total liabilities", year)
        assets = fact_ref(positions, "Balance Sheet", "Total assets", year)
        revenue = fact_ref(positions, "Income Statement", "Revenue", year)
        income = fact_ref(positions, "Income Statement", "Net income", year)
        prev = fact_ref(positions, "Balance Sheet", "Total assets", year - 1)
        if not prev and year == first_year and prior_assets is not None:
            prev = f"'Financial Ratios'!$H$2"
        notes = []
        if ca and cl:
            sheet.cell(row, 2, f'=IFERROR({ca}/{cl},"")')
        else:
            notes.append("Current ratio missing source facts")
        if liab and assets:
            sheet.cell(row, 3, f'=IFERROR({liab}/{assets},"")')
        else:
            notes.append("Liabilities/assets missing source facts")
        if revenue and assets and prev:
            sheet.cell(row, 4,
                       f'=IFERROR({revenue}/AVERAGE({assets},{prev}),"")')
        else:
            notes.append("Asset turnover needs revenue and two asset balances")
        if income and revenue:
            sheet.cell(row, 5, f'=IFERROR({income}/{revenue},"")')
        else:
            notes.append("Net profit margin missing source facts")
        sheet.cell(row, 6, "; ".join(notes) if notes else "Ratio inputs available")
        for col in (2, 4):
            sheet.cell(row, col).number_format = '0.00"x"'
        for col in (3, 5):
            sheet.cell(row, col).number_format = '0.00%'
    style_sheet(sheet, widths=[16, 20, 24, 20, 23, 48, 36, 24])
    sheet.auto_filter.ref = f"A1:F{sheet.max_row}"
    sheet["G2"].font = Font(bold=True, color=NAVY)
    sheet.sheet_properties.tabColor = "3B82F6"


def add_sentiment(wb, sentiment):
    """Write seven category counts, clean word totals, and formula P/N ratios."""
    sheet = wb.create_sheet("Sentiment Analysis")
    sheet.append(["Year", "Cleaned File", "Total Words", *CATEGORIES,
                  "P/N Ratio", "Extraction Status"])
    for year in sorted(sentiment):
        filename, count, categories, status = sentiment[year]
        row = sheet.max_row + 1
        sheet.append([year, filename, count,
                      *(categories.get(cat) if categories else None
                        for cat in CATEGORIES),
                      None, status])
        if categories:
            sheet.cell(row, 11, f'=IF(D{row}=0,"N/A",E{row}/D{row})')
            sheet.cell(row, 11).number_format = "0.000"
    style_sheet(sheet, widths=[12, 59, 18] + [17]*7 + [17, 42])
    sheet.sheet_properties.tabColor = "0F766E"
    return {sheet.cell(row, 1).value: row
            for row in range(2, sheet.max_row + 1)}


def add_prices(wb, filings, prices, ticker):
    """Write market prices with exact balance sheet and trading dates."""
    sheet = wb.create_sheet("Stock Prices")
    sheet.append(["Fiscal Year", "Balance Sheet Date", "Trading Date",
                  "Close (USD)", "Ticker", "Price Basis"])
    for filing in sorted(filings, key=lambda f: f["reportDate"]):
        year = int(filing_year(filing))
        balance = date.fromisoformat(filing["reportDate"])
        result = prices.get(year)
        trading, close = result[1] if result else (None, None)
        sheet.append([year, balance, trading, close, ticker,
                      "Unadjusted daily close (Yahoo Finance)"])
        row = sheet.max_row
        for col in (2, 3):
            sheet.cell(row, col).number_format = "yyyy-mm-dd"
        sheet.cell(row, 4).number_format = '$#,##0.00'
    style_sheet(sheet, widths=[17, 24, 21, 21, 15, 45])
    sheet.sheet_properties.tabColor = "B45309"
    return {sheet.cell(row, 1).value: row
            for row in range(2, sheet.max_row + 1)}


def add_final(wb, filings, sentiment_rows, price_rows):
    """Produce chronological chart data and two clearly labeled y-axes."""
    sheet = wb.create_sheet("Final Analysis")
    sheet.append(["Fiscal Year", "Stock Close (USD)", "P/N Ratio",
                  "Trading Date", "Notes"])
    for filing in sorted(filings, key=lambda f: f["reportDate"]):
        year = int(filing_year(filing))
        row = sheet.max_row + 1
        sheet.cell(row, 1, year)
        pr = price_rows[year]
        sr = sentiment_rows[year]
        if wb["Stock Prices"].cell(pr, 4).value is not None:
            sheet.cell(row, 2, f"='Stock Prices'!D{pr}")
            sheet.cell(row, 4, f"='Stock Prices'!C{pr}")
        if wb["Sentiment Analysis"].cell(sr, 4).value is not None:
            sheet.cell(row, 3, f"=IF('Sentiment Analysis'!D{sr}=0,NA(),"
                       f"'Sentiment Analysis'!E{sr}/'Sentiment Analysis'!D{sr})")
        sheet.cell(row, 5, "Compare trends; correlation is not causation")
        sheet.cell(row, 2).number_format = '$#,##0.00'
        sheet.cell(row, 3).number_format = "0.000"
        sheet.cell(row, 4).number_format = "yyyy-mm-dd"
    style_sheet(sheet, widths=[17, 23, 19, 21, 57])
    sheet.sheet_properties.tabColor = "7C3AED"
    if sheet.max_row < 3:
        sheet["G2"] = "Chart needs at least two fiscal years."
        return
    # Separate charts avoid unreliable dual-axis rendering in Excel.
    # Each chart has one measure, its own y-axis, and fiscal-year labels.
    categories = Reference(sheet, min_col=1, min_row=2,
                           max_row=sheet.max_row)

    def add_trend_chart(column, title, axis_label, number_format,
                        line_color, anchor):
        """Build a single-axis chart that scales to the selected year count."""
        chart = LineChart()
        chart.title = None  # Heading is in worksheet; avoid overlapping point labels.
        chart.style = 13
        chart.width = 23
        chart.height = 10.2
        chart.legend = None  # One series; a legend would be redundant.
        chart.display_blanks = "gap"
        chart.add_data(Reference(sheet, min_col=column, min_row=1,
                                 max_row=sheet.max_row),
                       titles_from_data=True)
        chart.set_categories(categories)
        chart.x_axis.title = None  # Year shown at each plotted point
        chart.x_axis.tickLblPos = "nextTo"
        chart.x_axis.tickMarkSkip = 1
        chart.x_axis.tickLblSkip = 1
        chart.y_axis.title = None  # Value shown at each plotted point
        chart.y_axis.numFmt = number_format
        chart.y_axis.tickLblPos = "nextTo"
        chart.y_axis.crosses = "min"
        chart.x_axis.crosses = "min"
        chart.y_axis.majorGridlines = None
        series = chart.series[0]
        series.graphicalProperties.line.solidFill = line_color
        series.graphicalProperties.line.width = 22000
        series.marker.symbol = "circle"
        series.marker.size = 6
        series.marker.graphicalProperties.solidFill = line_color
        series.marker.graphicalProperties.line.solidFill = line_color
        series.smooth = False
        # Excel can suppress tick labels depending on the chart engine.
        # Point labels show BOTH fiscal year and measure, independently
        # of axis-tick rendering, making the chart self-contained.
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showCatName = True
        chart.dataLabels.showVal = True
        chart.dataLabels.showSerName = False
        chart.dataLabels.separator = "\n"
        chart.dataLabels.dLblPos = "t"
        chart.dataLabels.numFmt = number_format
        sheet.add_chart(chart, anchor)

    sheet["G2"] = "FISCAL YEAR-END STOCK CLOSING PRICE"
    sheet["G2"].font = Font(name="Aptos", size=13, bold=True, color=NAVY)
    add_trend_chart(2, "Fiscal Year-End Stock Closing Price",
                    "Closing price (USD)", '$0', "245A81", "G3")
    sheet["G23"] = "MD&A POSITIVE-TO-NEGATIVE SENTIMENT RATIO"
    sheet["G23"].font = Font(name="Aptos", size=13, bold=True, color=NAVY)
    add_trend_chart(3, "MD&A Positive-to-Negative Sentiment Ratio",
                    "Positive / negative ratio", '0.00', "D18A35", "G24")
    sheet["G22"] = "Data labels show fiscal year and value; axes may vary by Excel version."
    sheet["G47"] = "Stock prices are unadjusted and may reflect stock splits."
    sheet["G48"] = "Sentiment scores require manual MD&A boundary validation."
    sheet["G49"] = "Compare trends cautiously; correlation is not causation."


def add_log(wb, log):
    """Record sources and MD&A quality-control flags for auditing."""
    sheet = wb.create_sheet("Extraction Log")
    sheet.append(["Fiscal Year", "Accession", "SEC Filing URL", "HTML File",
                  "MD&A Text", "Cleaned MD&A", "Status"])
    for entry in sorted(log, key=lambda x: x[0]):
        sheet.append(entry)
    style_sheet(sheet, widths=[17, 27, 73, 68, 68, 68, 63])
    sheet.sheet_properties.tabColor = "64748B"


def add_data_quality(wb, positions, filings, facts):
    """Show independently computed balance-sheet and income-statement checks.

    A blank means the required SEC facts were not available. The checks do not
    assume a missing number is zero.
    """
    sheet = wb.create_sheet("Data Quality")
    sheet.append(["Fiscal Year", "Assets - Liabilities - Equity (USD)",
                  "Balance Sheet Check", "Revenue - Cost - Gross Profit (USD)",
                  "Gross Profit Check", "Missing Reported Items", "Derived Items"])
    for filing in sorted(filings, key=lambda f: f["reportDate"]):
        year = int(filing_year(filing))
        def value(statement, label):
            entry = positions.get((statement, label, year))
            return entry[2] if entry is not None else None
        assets = value("Balance Sheet", "Total assets")
        liabilities = value("Balance Sheet", "Total liabilities")
        equity = value("Balance Sheet", "Stockholders' equity")
        revenue = value("Income Statement", "Revenue")
        cost = value("Income Statement", "Cost of revenue")
        gross = value("Income Statement", "Gross profit")
        bs_delta = (assets - liabilities - equity if all(
            isinstance(v, (int, float)) for v in (assets, liabilities, equity)) else None)
        gp_delta = (revenue - cost - gross if all(
            isinstance(v, (int, float)) for v in (revenue, cost, gross)) else None)
        missing = [f"{statement}: {label}" for statement, labels in STATEMENT_TAGS.items()
                   for label in labels if value(statement, label) is None]
        derived = []
        if value("Income Statement", "Operating expenses") is not None:
            original = find_fact(facts, STATEMENT_TAGS["Income Statement"]["Operating expenses"], filing, False)
            if original is None:
                derived.append("Operating expenses = gross profit - operating income")
        sheet.append([year, bs_delta, "PASS" if bs_delta == 0 else
                      ("MISSING" if bs_delta is None else "REVIEW"),
                      gp_delta, "PASS" if gp_delta == 0 else
                      ("MISSING" if gp_delta is None else "REVIEW"),
                      "; ".join(missing) if missing else "None",
                      "; ".join(derived) if derived else "None"])
    style_sheet(sheet, widths=[17, 36, 24, 42, 24, 100, 70])
    for row in sheet.iter_rows(min_row=2):
        for col in (2, 4):
            row[col-1].number_format = '#,##0;[Red](#,##0);–'
    sheet.sheet_properties.tabColor = "B91C1C"


def build_final_workbook(facts, filings, company, ticker, sentiment,
                         prices, log, output_path):
    """Generate the complete formatted Excel deliverable without overwriting."""
    # Choose an unused name again at save time, protecting existing results.
    base = output_path
    version = 2
    while output_path.exists():
        output_path = base.with_name(f"{base.stem}_v{version}{base.suffix}")
        version += 1
    print(f"Saving Excel workbook to: {output_path}")
    wb = Workbook()
    wb.remove(wb.active)
    positions = add_financial_sheets(wb, facts, filings, company, ticker)
    add_ratios(wb, positions, filings, facts)
    sentiment_rows = add_sentiment(wb, sentiment)
    price_rows = add_prices(wb, filings, prices, ticker)
    add_final(wb, filings, sentiment_rows, price_rows)
    add_log(wb, log)
    add_data_quality(wb, positions, filings, facts)
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)
    wb.active = wb.sheetnames.index("Final Analysis")
    wb.save(output_path)
    return output_path


def main():
    """Run all stages using the requested ticker and filing count."""
    ticker, count, root, identity, dictionary_path = prompt_settings()
    dictionary = read_dictionary(dictionary_path)
    client = SecClient(identity)
    cik, company = lookup_company(client, ticker)
    print("SCRIPT VERSION: FINAL_SAVE_FIXED_2026_11_04")
    print(f"Found {company} (CIK {cik}).")
    filings = recent_10ks(client, cik, count)
    if not filings:
        raise RuntimeError("No 10-K filings found for this ticker")
    if len(filings) < count:
        print(f"WARNING: Only {len(filings)} of {count} filings were found.")
    # Reject duplicate fiscal years to avoid misleading repeated columns.
    years = [int(filing_year(f)) for f in filings]
    if len(set(years)) != len(years):
        raise ValueError("Duplicate fiscal years detected; manual review required")
    company_dir = root / ticker
    raw_dir = company_dir / "10K_HTML"
    text_dir = company_dir / "MDA_Text"
    clean_dir = company_dir / "MDA_Cleaned"
    for directory in (raw_dir, text_dir, clean_dir):
        directory.mkdir(parents=True, exist_ok=True)
    sentiment = {}
    log = []
    for filing in filings:
        year = int(filing_year(filing))
        accession = filing["accessionNumber"]
        print(f"Processing {year}: {accession}...")
        url = f"{ARCHIVES_URL}/{cik}/{accession.replace('-', '')}/{filing['primaryDocument']}"
        html_path = ""
        mda_path = ""
        cleaned_path = ""
        status = ""
        try:
            downloaded, url = download_filing(client, cik, ticker, filing, raw_dir)
            html_path = str(downloaded)
            mda = save_mda(downloaded, ticker, filing, text_dir)
            if not mda:
                raise ValueError("MD&A extraction failed; boundaries not found")
            mda_path = str(mda)
            cleaned, total, categories = score_text(
                mda.read_text(encoding="utf-8"), dictionary)
            cleaned_file = clean_dir / (mda.stem + "_cleaned.txt")
            cleaned_file.write_text(cleaned + "\n", encoding="utf-8")
            cleaned_path = str(cleaned_file)
            status = "Extracted; REVIEW ITEM 7 BOUNDARIES"
            sentiment[year] = (cleaned_file.name, total, categories, status)
            neg = categories["Negative"]
            ratio = categories["Positive"] / neg if neg else None
            print(f"  Words={total:,}; Positive={categories['Positive']}; "
                  f"Negative={neg}; P/N={ratio if ratio is not None else 'N/A'}")
        except (RuntimeError, OSError, ValueError, UnicodeError) as error:
            status = f"FAILED: {error}"
            print(f"  WARNING: {status}")
            sentiment[year] = ("", None, None, status)
        log.append([year, accession, url, html_path, mda_path,
                    cleaned_path, status])
    print("Retrieving SEC XBRL financial facts...")
    facts = client.get(FACTS_URL.format(cik=f"{cik:010d}"))
    print("Retrieving historical closing prices...")
    try:
        prices = fetch_prices(ticker, filings)
    except (RuntimeError, ValueError, OSError, requests.RequestException) as error:
        print(f"WARNING: Price retrieval failed: {error}")
        prices = {}
    base = company_dir / f"{ticker}_Financial_Analysis_{min(years)}-{max(years)}.xlsx"
    output = base
    revision = 2
    while output.exists():
        output = base.with_name(f"{base.stem}_v{revision}{base.suffix}")
        revision += 1
    if output != base:
        print(f"Existing workbook preserved; saving new version as {output.name}")
    saved = build_final_workbook(facts, filings, company, ticker,
                                 sentiment, prices, log, output)
    print(f"\nSaved workbook: {saved}")
    print(f"Requested {count}; retrieved {len(filings)} filing(s).")
    print("QUALITY WARNING: MD&A boundaries require review against original 10-Ks.")
    print("Review Data Quality for missing/derived XBRL fields before submission.")
    print("Open in Excel to recalculate formulas and display the chart.")
    print("Do not treat unverified missing data as zero.")


if __name__ == "__main__":
    try:
        main()
    except (RuntimeError, ValueError, OSError, KeyError) as error:
        print(f"ERROR: {error}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("Cancelled by user.")
        sys.exit(130)
