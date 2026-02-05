import requests
import os
import json
import pandas as pd
import webbrowser
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys

# =====================================================
# CONFIG
# =====================================================
RUN_MODE = os.environ.get("RUN_MODE", "interactive")  # interactive | scheduled

BASE_PATH = os.path.dirname(os.path.abspath(__file__))

EXCEL_ROOT = os.path.join(BASE_PATH, "axis_monthly_portfolios")
CSV_ROOT = os.path.join(BASE_PATH, "axis_monthly_csv")

EXCEL_STATE_FILE = os.path.join(BASE_PATH, "downloaded_files.json")
CSV_STATE_FILE = os.path.join(BASE_PATH, "generated_csv_files.json")
PENDING_FILE = os.path.join(BASE_PATH, "pending_months.json")

BASE_URL = "https://www.axismf.com/cms/sites/default/files/Statutory/"
AMC_NAME = "Axis Mutual Fund"

HISTORICAL_START = datetime(2021, 1, 1)

os.makedirs(EXCEL_ROOT, exist_ok=True)
os.makedirs(CSV_ROOT, exist_ok=True)

# =====================================================
# STATE HELPERS
# =====================================================
def load_state(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()


def save_state(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(sorted(data), f, indent=2)


# =====================================================
# DATA HELPERS
# =====================================================
def clean_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.lower()
        .str.replace("\n", " ")
        .str.replace("%", "percent")
        .str.replace(".", "", regex=False)
        .str.strip()
    )
    return df


def excel_to_csv(excel_path, csv_path):
    xls = pd.ExcelFile(excel_path)
    all_data = []

    for sheet in xls.sheet_names:
        if sheet.lower() == "index":
            continue

        raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        scheme_name = str(raw.iloc[0, 1])

        df = pd.read_excel(xls, sheet_name=sheet, header=3)
        df = clean_columns(df)

        column_map = {
            "name of the instrument": "instrument_name",
            "isin": "isin",
            "industry / rating": "industry_or_rating",
            "industry rating": "industry_or_rating",
            "quantity": "quantity",
            "market value rs in lakhs": "market_value",
            "market fair value rs in lakhs": "market_value",
            "percent to net assets": "percentage_of_portfolio",
        }

        df = df.rename(columns=column_map)
        df = df[[c for c in column_map.values() if c in df.columns]]

        if "isin" in df.columns:
            df = df[df["isin"].notna()]

        df = df.loc[:, ~df.columns.duplicated()]

        df["amc_name"] = AMC_NAME
        df["scheme_name"] = scheme_name
        df["reporting_date"] = datetime.today().strftime("%Y-%m-%d")

        all_data.append(df)

    if not all_data:
        return False

    final_df = pd.concat(all_data, ignore_index=True)
    final_df.to_csv(csv_path, index=False)
    return True


# =====================================================
# MANUAL FALLBACK
# =====================================================
def manual_excel_fallback(month_year, csv_dir, pending_months):
    if RUN_MODE == "scheduled":
        print(f"⚠️ MANUAL ACTION REQUIRED: {month_year}")
        return

    print("\n⚠️ AUTOMATIC DOWNLOAD FAILED")
    print(f"📅 Month: {month_year}")
    print("🌐 Opening Axis MF Statutory Disclosure page...")
    print("👉 Download the Excel file manually.")
    print("👉 Paste FULL path to .xlsx file (Enter to skip)\n")

    webbrowser.open("https://www.axismf.com/statutory-disclosures")
    excel_path = input("📂 Excel path: ").strip()
    excel_path = excel_path.strip('"').strip("'")


    if not excel_path:
        print("⏭️ Skipped")
        return
    if not os.path.exists(excel_path) or os.path.isdir(excel_path):
        print("❌ Invalid file path")
        return
    if not excel_path.lower().endswith((".xlsx", ".xls")):
        print("❌ Not an Excel file")
        return

    csv_name = os.path.basename(excel_path).rsplit(".", 1)[0] + ".csv"
    csv_path = os.path.join(csv_dir, csv_name)

    try:
        excel_to_csv(excel_path, csv_path)
        pending_months.discard(month_year)
        print(f"📄 CSV created: {csv_path}")
    except Exception as e:
        print(f"❌ Conversion failed: {e}")


# =====================================================
# LOAD STATE
# =====================================================
downloaded_excels = load_state(EXCEL_STATE_FILE)
generated_csvs = load_state(CSV_STATE_FILE)
pending_months = load_state(PENDING_FILE)

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0"})

today = datetime.today().replace(day=1)

if RUN_MODE == "interactive":
    start_date = HISTORICAL_START
else:
    start_date = today - relativedelta(months=3)

current = today

# =====================================================
# MAIN LOOP
# =====================================================
while current >= start_date:
    last_day = current + relativedelta(months=1) - relativedelta(days=1)
    month_year = last_day.strftime("%B %Y")
    year = str(last_day.year)

    print(f"\n📅 Processing: {month_year}")

    excel_dir = os.path.join(EXCEL_ROOT, year)
    csv_dir = os.path.join(CSV_ROOT, year)
    os.makedirs(excel_dir, exist_ok=True)
    os.makedirs(csv_dir, exist_ok=True)

    excel_name = f"Monthly Portfolio-{last_day.strftime('%d %m %y')}.xlsx"
    csv_name = excel_name.replace(".xlsx", ".csv")

    excel_path = os.path.join(excel_dir, excel_name)
    csv_path = os.path.join(csv_dir, csv_name)

    url = BASE_URL + excel_name.replace(" ", "%20")

    # -------- EXCEL --------
    if os.path.exists(excel_path):
        print(f"↩️ Excel exists: {excel_name}")
        downloaded_excels.add(excel_name)
        pending_months.discard(month_year)

    else:
        r = session.get(url, timeout=20)
        if r.status_code == 200 and "application" in r.headers.get("Content-Type", ""):
            with open(excel_path, "wb") as f:
                f.write(r.content)
            downloaded_excels.add(excel_name)
            pending_months.discard(month_year)
            print(f"📥 Excel downloaded: {excel_name}")

        else:
            pending_months.add(month_year)
            print(f"📌 Pending: {month_year}")
            manual_excel_fallback(month_year, csv_dir, pending_months)
            current -= relativedelta(months=1)
            continue

    # -------- CSV --------
    if csv_name in generated_csvs and os.path.exists(csv_path):
        print(f"↩️ CSV exists: {csv_name}")
    else:
        if excel_to_csv(excel_path, csv_path):
            generated_csvs.add(csv_name)
            print(f"📄 CSV created: {csv_name}")

    current -= relativedelta(months=1)

# =====================================================
# SAVE STATE & EXIT
# =====================================================
save_state(EXCEL_STATE_FILE, downloaded_excels)
save_state(CSV_STATE_FILE, generated_csvs)
save_state(PENDING_FILE, pending_months)

sys.exit(0)
