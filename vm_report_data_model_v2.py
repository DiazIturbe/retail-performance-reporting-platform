"""
VM Weekly Report Data Model Merger v3

Reads the Tableau CSV exports for the weekly VM report and builds a clean,
report-ready data model. This version includes Traffic By The Hour exported as
Table.csv and calculates Footfall / Conversion from that file.

Required files in the working folder:
- LFL.csv
- Group by Site.xlsx or Group by Site.csv (department summary)
- D Net Sales.csv (subdepartment sales + budget)
- Footwear.csv
- Apparel.csv
- Accessories.csv
- Table.csv  (Tableau report: Traffic By The Hour)
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import re
from typing import Dict, List, Tuple

import pandas as pd


INPUT_FILES = {
    "lfl": "LFL.csv",
    "group_by_site": "Group by Site.xlsx",
    "d_net_sales": "D Net Sales.csv",
    "footwear": "Footwear.csv",
    "apparel": "Apparel.csv",
    "accessories": "Accessories.csv",
    "traffic_by_hour": "Table.csv",
    "option_list": "Option List.csv",
}

SUBDEPARTMENT_ORDER = [
    ("Apparel", "Mens Apparel"),
    ("Apparel", "Womens Apparel"),
    ("Apparel", "Junior Apparel"),
    ("Apparel", "Childrens Apparel"),
    ("Footwear", "Mens Footwear"),
    ("Footwear", "Womens Footwear"),
    ("Footwear", "Junior Footwear"),
    ("Footwear", "Childrens Footwear"),
    ("Footwear", "Infants Footwear"),
    ("Accessories", "Clothing Accessories"),
    ("Accessories", "Other Accessories"),
]

DEPARTMENT_ORDER = {"Footwear": 1, "Apparel": 2, "Accessories": 3}
SUBDEPT_SORT_ORDER = {sub: i for i, (_, sub) in enumerate(SUBDEPARTMENT_ORDER)}


def read_tableau_csv(path: str | Path) -> pd.DataFrame:
    """
    Read a Tableau export from CSV or Excel.

    Supports:
    - UTF-16 Tableau CSV exports
    - UTF-8 CSV exports
    - XLSX exports from newer Tableau downloads

    Important Streamlit note:
    The uploaded file may be saved using the internal expected filename.
    For example, a CSV upload may temporarily be saved as Group by Site.xlsx.
    Therefore, if Excel reading fails, this function falls back to CSV parsing.
    """
    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(f"Missing required input file: {path.name}")

    # Try Excel first only when the filename suggests Excel.
    # If it fails, continue to CSV fallback instead of stopping.
    if path.suffix.lower() in [".xlsx", ".xls"]:
        try:
            return pd.read_excel(path, engine="openpyxl")
        except Exception as exc:
            print(f"Excel read failed for {path.name}: {exc}")

    last_error = None

    for encoding in ["utf-16", "utf-8-sig", "utf-8", "latin1"]:
        for sep in ["\t", ",", ";"]:
            try:
                df = pd.read_csv(path, encoding=encoding, sep=sep)

                if len(df.columns) > 1:
                    return df

            except Exception as exc:
                last_error = exc

    raise ValueError(
        f"Could not read file '{path.name}'. Please verify the Tableau export format. "
        f"Last error: {last_error}"
    )

def money_to_float(value) -> float:
    """Convert Tableau money values to float.

    Handles both North American format ($64,396) and Tableau/European-style
    exports where dots are thousands separators and commas are decimals
    ($64.396 means 64,396, not 64.396).
    """
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace("\xa0", " ")
    if text in ["", "nan", "NaN", "None", "-"]:
        return 0.0

    neg = text.startswith("-") or (text.startswith("(") and text.endswith(")"))
    text = (
        text.replace("$", "")
        .replace("%", "")
        .replace("(", "")
        .replace(")", "")
        .replace("-", "")
        .replace(" ", "")
        .strip()
    )

    # Keep only digits and separators.
    text = re.sub(r"[^0-9.,]", "", text)
    if text == "":
        return 0.0

    if "." in text and "," in text:
        # Last separator usually identifies the decimal separator.
        # 1.234,56 -> 1234.56 ; 1,234.56 -> 1234.56
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "." in text:
        # Tableau export: 64.396, 5.399, 1.234.567 = thousands groups.
        if re.fullmatch(r"\d{1,3}(\.\d{3})+", text):
            text = text.replace(".", "")
        # Otherwise keep decimal dot, e.g. 64.39
    elif "," in text:
        # 64,396 = thousands; 64,39 or 64,3 = decimal comma.
        if re.fullmatch(r"\d{1,3}(,\d{3})+", text):
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")

    try:
        number = float(text)
    except ValueError:
        return 0.0
    return -number if neg else number


def pct_to_float(value) -> float:
    """Convert Tableau percentages to decimal proportions.

    Handles 6.7%, 6,7%, 100,00%, Excel numeric 0.067, and whole-number 6.7.
    """
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        value = float(value)
        return value if abs(value) <= 1 else value / 100

    text = str(value).strip().replace("\xa0", " ").replace("%", "").replace(" ", "")
    if text in ["", "nan", "NaN", "None", "-"]:
        return 0.0

    text = re.sub(r"[^0-9.,\-]", "", text)
    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        # Percent exports normally use comma as decimal: 47,7 -> 47.7
        text = text.replace(",", ".")

    try:
        number = float(text)
        return number / 100
    except ValueError:
        return 0.0

def num_to_float(value) -> float:
    if pd.isna(value):
        return 0.0
    text = str(value).strip().replace("\xa0", " ")
    if text in ["", "nan", "NaN", "None", "-"]:
        return 0.0
    # Tableau exports may use either 1,084 as thousands or 7,0 as decimal.
    if "," in text and "." not in text:
        parts = text.split(",")
        if len(parts[-1]) == 3 and all(part.isdigit() for part in parts):
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
    else:
        text = text.replace(",", "")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in ["", "-", "."]:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def safe_divide(numerator: float, denominator: float, default: float = 0.0) -> float:
    if denominator in (0, None) or pd.isna(denominator):
        return default
    return numerator / denominator


def clean_lfl(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy().rename(columns={"Unnamed: 0": "business", "Unnamed: 1": "site"})

    if "site" in out.columns:
        out = out[out["site"].astype(str).str.contains("3012", na=False)].copy()

    if out.empty:
        out = df.tail(1).copy().rename(columns={"Unnamed: 0": "business", "Unnamed: 1": "site"})

    rename = {
        "Net Sales": "net_sales",
        "Net Sales $": "net_sales",

        "Previous Year Net Sales": "py_net_sales",
        "PY Net Sales": "py_net_sales",

        "Net Sales Previous Year VAR": "net_sales_py_var",
        "Sales LFL": "sales_lfl_pct",

        "Transactions": "transactions_lfl_source",
        "Transactions ": "transactions_lfl_source",

        "Previous Year Transactions": "py_transactions",
        "LY Transactions ": "py_transactions",

        "Transactions Previous Year VAR": "transactions_py_var",
        "Transactions LFL": "transactions_lfl_pct",

        "IPC": "ipc",
        "UPT (LFL)": "ipc",

        "Previous Year IPC": "py_ipc",
        "LY UPT": "py_ipc",

        "IPC Previous Year VAR": "ipc_py_var",
        "UPT LY VAR": "ipc_py_var",

        "IPC LFL": "ipc_lfl_pct",
        "UPT LFL": "ipc_lfl_pct",

        "Footfall (LFL)": "footfall",
        "FootFall LY": "py_footfall",
        "ATV (LFL)": "atv",
        "LY ATV": "py_atv",
        "Conversion (LFL)": "conversion_pct",
        "LY Conversion": "py_conversion_pct",
    }

    out = out.rename(columns=rename)

    required_cols = [
        "site",
        "net_sales",
        "py_net_sales",
        "net_sales_py_var",
        "sales_lfl_pct",
        "transactions_lfl_source",
        "py_transactions",
        "transactions_py_var",
        "transactions_lfl_pct",
        "ipc",
        "py_ipc",
        "ipc_py_var",
        "ipc_lfl_pct",
        "footfall",
        "py_footfall",
        "atv",
        "py_atv",
        "conversion_pct",
        "py_conversion_pct",
    ]

    for col in required_cols:
        if col not in out.columns:
            out[col] = 0

    for col in ["net_sales", "py_net_sales", "net_sales_py_var"]:
        out[col] = out[col].apply(money_to_float)

    for col in ["sales_lfl_pct", "transactions_lfl_pct", "ipc_lfl_pct", "conversion_pct", "py_conversion_pct"]:
        out[col] = out[col].apply(pct_to_float)

    for col in [
        "transactions_lfl_source",
        "py_transactions",
        "transactions_py_var",
        "ipc",
        "py_ipc",
        "ipc_py_var",
        "footfall",
        "py_footfall",
        "atv",
        "py_atv",
    ]:
        out[col] = out[col].apply(money_to_float if col in ["atv", "py_atv"] else num_to_float)

    return out[required_cols]

def clean_department_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Clean department-level summary from the new Group TY / Group by Site export.

    The new Tableau download has one row per group (Accessories, Apparel,
    Footwear) and already includes Net Sales, Budget CA, VAR, VAR %, Avg Sales
    Price, and Units. Older exports with site/banner columns are still partly
    supported.
    """
    out = df.copy().rename(columns={
        "Unnamed: 0": "group",
        "Unnamed: 1": "banner",
        "Unnamed: 2": "site",
        "Group": "group",
        "Net Sales": "net_sales",
        "% of Net Sales  ": "net_sales_mix_pct",
        "% of Net Sales ": "net_sales_mix_pct",
        "% of Net Sales": "net_sales_mix_pct",
        "% de total Net Sales junto con Group": "net_sales_mix_pct",
        "Margin $": "margin",
        "Margin %": "margin_pct",
        "Budget CA": "budget_group_by_site",
        "VAR (CA)": "budget_gap_group_by_site",
        "VAR % (CA)": "budget_gap_pct_group_by_site",
        "Avg Sales Price": "avg_sales_price",
        "Units": "units",
    })

    if "site" in out.columns:
        filtered = out[out["site"].astype(str).str.contains("3012", na=False)].copy()
        if not filtered.empty:
            out = filtered
    else:
        out["site"] = "3012"

    for required in [
        "group", "net_sales", "net_sales_mix_pct", "budget_group_by_site",
        "budget_gap_group_by_site", "budget_gap_pct_group_by_site",
        "avg_sales_price", "units"
    ]:
        if required not in out.columns:
            out[required] = 0

    if "margin" not in out.columns:
        out["margin"] = 0
    if "margin_pct" not in out.columns:
        out["margin_pct"] = 0

    for col in ["net_sales", "margin", "budget_group_by_site", "budget_gap_group_by_site", "avg_sales_price"]:
        out[col] = out[col].apply(money_to_float)
    for col in ["net_sales_mix_pct", "margin_pct", "budget_gap_pct_group_by_site"]:
        out[col] = out[col].apply(pct_to_float)
    out["units"] = out["units"].apply(num_to_float)

    out = out[out["group"].astype(str).isin(["Footwear", "Apparel", "Accessories"])]

    return out[[
        "group", "site", "net_sales", "net_sales_mix_pct", "margin", "margin_pct",
        "budget_group_by_site", "budget_gap_group_by_site", "budget_gap_pct_group_by_site",
        "avg_sales_price", "units"
    ]]


def clean_department_targets(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Clean subdepartment budgets from the new combined D Net Sales export.

    The previous workflow used a separate D Targets.csv. The new Tableau export
    combines Net Sales and Budget CA in D Net Sales.csv, so this function now
    reads budget fields from that same file.
    """
    out = df.copy().rename(columns={
        "Unnamed: 0": "group",
        "Unnamed: 1": "department",
        "Group": "group",
        "Department Description": "department",
        "Budget": "budget",
        "Budget CA": "budget",
        "% de total Budget junto con Department Description": "budget_mix_pct",
        "% de total Budget CA junto con Group, DEPARTMENT_DESC": "budget_mix_pct",
        "VAR (CA)": "budget_gap_export",
        "VAR % (CA)": "budget_gap_pct_export",
    })

    for required in ["group", "department", "budget", "budget_mix_pct", "budget_gap_export", "budget_gap_pct_export"]:
        if required not in out.columns:
            out[required] = 0

    out["budget"] = out["budget"].apply(money_to_float)
    out["budget_gap_export"] = out["budget_gap_export"].apply(money_to_float)
    out["budget_gap_pct_export"] = out["budget_gap_pct_export"].apply(pct_to_float)
    out["budget_mix_pct"] = out["budget_mix_pct"].apply(pct_to_float)

    totals = out[out["department"].astype(str).str.lower() == "total"].copy()
    detail = out[out["department"].astype(str).str.lower() != "total"].copy()
    return detail, totals


def clean_department_sales(df: pd.DataFrame) -> pd.DataFrame:
    """Clean subdepartment sales from D Net Sales.csv.

    Supports both older and newer Tableau headers.
    """
    out = df.copy().rename(columns={
        "Unnamed: 0": "group",
        "Unnamed: 1": "department",
        "Unnamed: 2": "department",
        "Net Sales": "net_sales",
        "% of Net Sales ": "net_sales_mix_within_group_pct",
        "% of Net Sales": "net_sales_mix_within_group_pct",
        "% de total Net Sales junto con Group, DEPARTMENT_DESC": "net_sales_mix_within_group_pct",
        "Units": "units",
    })

    for required in ["group", "department", "net_sales", "net_sales_mix_within_group_pct", "units"]:
        if required not in out.columns:
            out[required] = 0

    out = out[out["department"].astype(str).str.lower() != "total"].copy()
    out["net_sales"] = out["net_sales"].apply(money_to_float)
    out["net_sales_mix_within_group_pct"] = out["net_sales_mix_within_group_pct"].apply(pct_to_float)
    out["units"] = out["units"].apply(num_to_float)

    out = out[out["group"].astype(str).isin(["Footwear", "Apparel", "Accessories"])]

    return out[["group", "department", "net_sales", "net_sales_mix_within_group_pct", "units"]]


def clean_brand_file(df: pd.DataFrame, group_name: str) -> pd.DataFrame:
    out = df.copy().rename(columns={
        "Unnamed: 0": "department",
        "Unnamed: 1": "brand",
        "Net Sales": "net_sales",
        "% of Net Sales": "net_sales_mix_pct",
        "% of Gender": "department_brand_mix_pct",
        "Margin": "margin",
        "Margin %": "margin_pct",
        "Units Sold": "units",
        "% of Units Sold": "units_mix_pct",
    })
    out = out[out["department"].astype(str).str.lower() != "business total"].copy()
    out = out[out["brand"].astype(str).str.lower() != "total"].copy()
    out = out[out["brand"].astype(str).str.strip() != ""].copy()
    out["group"] = group_name
    for col in ["net_sales", "margin"]:
        out[col] = out[col].apply(money_to_float)
    for col in ["net_sales_mix_pct", "department_brand_mix_pct", "margin_pct", "units_mix_pct"]:
        out[col] = out[col].apply(pct_to_float)
    out["units"] = out["units"].apply(num_to_float)
    return out[[
        "group", "department", "brand", "net_sales", "net_sales_mix_pct",
        "department_brand_mix_pct", "margin", "margin_pct", "units", "units_mix_pct"
    ]]



def clean_option_list(df: pd.DataFrame) -> pd.DataFrame:
    """Clean Tableau Option List and aggregate duplicate style rows."""
    out = df.copy().rename(columns={
        "Style Id": "style_id", "Style No": "style_no",
        "Department Desc": "department", "Brand Desc": "brand",
        "Style Desc": "product", "Net Sales": "net_sales",
        "% of Net Sales": "sales_share_pct", "Margin": "margin",
        "Margin %": "margin_pct", "Units Sold": "units_sold",
        "% of Units Sold": "units_share_pct", "Units on Hand": "units_on_hand",
    })
    required = ["style_id", "style_no", "department", "brand", "product", "net_sales", "units_sold", "units_on_hand"]
    for col in required:
        if col not in out.columns:
            out[col] = "" if col in ["style_id", "style_no", "department", "brand", "product"] else 0.0
    for col in ["net_sales", "margin"]:
        if col in out.columns: out[col] = out[col].apply(money_to_float)
    for col in ["sales_share_pct", "margin_pct", "units_share_pct"]:
        if col in out.columns: out[col] = out[col].apply(pct_to_float)
    for col in ["units_sold", "units_on_hand"]:
        out[col] = out[col].apply(num_to_float)
    for col in ["style_id", "style_no", "department", "brand", "product"]:
        out[col] = out[col].fillna("").astype(str).str.strip()
    out = out[(out["product"] != "") & (out["net_sales"] > 0)].copy()
    out["group"] = out["department"].map(dict((sub, group) for group, sub in SUBDEPARTMENT_ORDER)).fillna("")
    out = out[out["group"] != ""].copy()
    keys = ["style_id", "style_no", "group", "department", "brand", "product"]
    agg = {"net_sales":"sum", "units_sold":"sum", "units_on_hand":"max"}
    if "margin" in out.columns: agg["margin"] = "sum"
    detail = out.groupby(keys, as_index=False).agg(agg)
    detail["atv_per_unit"] = [safe_divide(s,u) for s,u in zip(detail["net_sales"], detail["units_sold"])]
    total = detail["net_sales"].sum()
    detail["share_of_store_sales_pct"] = detail["net_sales"].apply(lambda x: safe_divide(x,total))
    return detail.sort_values("net_sales", ascending=False).reset_index(drop=True)


def build_top_sellers(option_detail: pd.DataFrame, global_n: int = 5, group_n: int = 3, sub_n: int = 3):
    def ranked(df, by, n):
        if df.empty: return df.copy()
        frames=[]
        for key, g in df.groupby(by, dropna=False):
            x=g.sort_values(["net_sales","units_sold"], ascending=[False,False]).head(n).copy()
            x["rank"] = range(1, len(x)+1)
            frames.append(x)
        return pd.concat(frames, ignore_index=True) if frames else df.head(0).copy()
    global_df=option_detail.sort_values(["net_sales","units_sold"], ascending=[False,False]).head(global_n).copy()
    global_df["rank"] = range(1,len(global_df)+1)
    return global_df, ranked(option_detail,["group"],group_n), ranked(option_detail,["group","department"],sub_n)

def hour_sort_key(hour: str) -> int:
    match = re.match(r"(\d+)\s*([ap])", str(hour).lower())
    if not match:
        return 99
    h = int(match.group(1))
    ap = match.group(2)
    if ap == "p" and h != 12:
        h += 12
    if ap == "a" and h == 12:
        h = 0
    return h


def is_hour_label(value) -> bool:
    """Return True for Tableau hour labels such as '10 a. m.' or '4 p.\xa0m.'."""
    text = str(value).lower().replace("\xa0", " ")
    compact = re.sub(r"\s+", "", text)
    return bool(re.search(r"\d{1,2}[ap]\.?m\.?", compact))


def clean_traffic_by_hour(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Clean Tableau 'Traffic By The Hour' export named Table.csv.

    Returns:
    - hourly_traffic: one row per hour across the full report period
    - traffic_summary: footfall/conversion summary and peak-hour facts

    Important:
    Tableau exports hour labels like '4 a.\xa0m.' or '4 a. m.'. The previous parser
    looked only for the exact pattern 'a.m.' / 'p.m.', so it skipped every row and
    returned zero footfall and zero transactions. This parser normalizes spaces and
    non-breaking spaces before detecting hours.
    """
    raw = df.copy()
    rows = raw.fillna("").astype(str).values.tolist()
    columns = list(raw.columns)

    day_blocks: List[Tuple[str, int]] = []
    for idx in range(1, len(columns), 6):
        day = str(columns[idx]).split(".")[0].strip()
        if day and "Unnamed" not in day:
            day_blocks.append((day, idx))

    hourly = defaultdict(lambda: {"footfall": 0.0, "transactions": 0.0, "net_sales": 0.0})

    # Skip the metric-name row when present.
    detail_rows = rows[1:] if rows and "Footfall" in " ".join(rows[0]) else rows

    for row in detail_rows:
        if not row:
            continue

        hour = str(row[0]).strip()
        if not is_hour_label(hour):
            continue
        if any(term in hour.lower() for term in ["total", "general"]):
            continue

        for _, idx in day_blocks:
            if idx + 2 >= len(row):
                continue

            ff = num_to_float(row[idx])
            trans = num_to_float(row[idx + 1])
            sales = money_to_float(row[idx + 2])

            if ff == 0 and trans == 0 and sales == 0:
                continue

            hourly[hour]["footfall"] += ff
            hourly[hour]["transactions"] += trans
            hourly[hour]["net_sales"] += sales

    hourly_rows = []
    for hour, values in sorted(hourly.items(), key=lambda item: hour_sort_key(item[0])):
        ff = values["footfall"]
        trans = values["transactions"]
        sales = values["net_sales"]
        hourly_rows.append({
            "hour": hour,
            "footfall": ff,
            "transactions": trans,
            "net_sales": sales,
            "conversion_pct": safe_divide(trans, ff),
        })

    hourly_df = pd.DataFrame(hourly_rows)

    total_footfall = float(hourly_df["footfall"].sum()) if not hourly_df.empty else 0.0
    total_transactions = float(hourly_df["transactions"].sum()) if not hourly_df.empty else 0.0
    total_sales = float(hourly_df["net_sales"].sum()) if not hourly_df.empty else 0.0

    if hourly_df.empty:
        peaks = {
            "peak_traffic_hour": "",
            "peak_traffic_value": 0,
            "peak_transaction_hour": "",
            "peak_transaction_value": 0,
            "peak_sales_hour": "",
            "peak_sales_value": 0,
            "best_conversion_hour": "",
            "best_conversion_value": 0,
        }
    else:
        peak_traffic = hourly_df.sort_values("footfall", ascending=False).iloc[0]
        peak_trans = hourly_df.sort_values("transactions", ascending=False).iloc[0]
        peak_sales = hourly_df.sort_values("net_sales", ascending=False).iloc[0]
        eligible = hourly_df[(hourly_df["footfall"] >= 50) & (hourly_df["conversion_pct"] > 0)]
        peak_conv = eligible.sort_values("conversion_pct", ascending=False).iloc[0] if not eligible.empty else hourly_df.sort_values("conversion_pct", ascending=False).iloc[0]

        peaks = {
            "peak_traffic_hour": peak_traffic["hour"],
            "peak_traffic_value": peak_traffic["footfall"],
            "peak_transaction_hour": peak_trans["hour"],
            "peak_transaction_value": peak_trans["transactions"],
            "peak_sales_hour": peak_sales["hour"],
            "peak_sales_value": peak_sales["net_sales"],
            "best_conversion_hour": peak_conv["hour"],
            "best_conversion_value": peak_conv["conversion_pct"],
        }

    summary_df = pd.DataFrame([
        {"metric": "Total Footfall", "value": total_footfall, "source": "Table.csv", "notes": "Sum of hourly footfall"},
        {"metric": "Total Transactions", "value": total_transactions, "source": "Table.csv", "notes": "Sum of hourly transactions"},
        {"metric": "Conversion %", "value": safe_divide(total_transactions, total_footfall), "source": "Table.csv", "notes": "Transactions / Footfall"},
        {"metric": "Traffic Net Sales", "value": total_sales, "source": "Table.csv", "notes": "Hourly traffic sales total"},
        {"metric": "Peak Traffic Hour", "value": peaks.get("peak_traffic_hour", ""), "source": "Table.csv", "notes": f"{peaks.get('peak_traffic_value', 0):,.0f} footfall"},
        {"metric": "Peak Transaction Hour", "value": peaks.get("peak_transaction_hour", ""), "source": "Table.csv", "notes": f"{peaks.get('peak_transaction_value', 0):,.0f} transactions"},
        {"metric": "Peak Sales Hour", "value": peaks.get("peak_sales_hour", ""), "source": "Table.csv", "notes": f"${peaks.get('peak_sales_value', 0):,.0f}"},
        {"metric": "Best Conversion Hour", "value": peaks.get("best_conversion_hour", ""), "source": "Table.csv", "notes": f"{peaks.get('best_conversion_value', 0):.1%}"},
    ])

    return hourly_df, summary_df


def build_department_summary(executive_kpis: pd.DataFrame, dept_summary_raw: pd.DataFrame, dept_target_totals: pd.DataFrame) -> pd.DataFrame:
    sales_total = float(executive_kpis["net_sales"].iloc[0]) if not executive_kpis.empty else dept_summary_raw["net_sales"].sum()
    rows = []
    for _, row in dept_summary_raw.iterrows():
        group = row["group"]
        target = dept_target_totals[dept_target_totals["group"] == group]
        budget = float(target["budget"].sum()) if not target.empty else float(row.get("budget_group_by_site", 0))
        gap = float(row["net_sales"]) - budget
        rows.append({
            "group": group,
            "site": row.get("site", ""),
            "net_sales": float(row["net_sales"]),
            "budget": budget,
            "budget_gap": gap,
            "budget_gap_pct": safe_divide(gap, budget),
            "net_sales_mix_pct": safe_divide(float(row["net_sales"]), sales_total),
            "margin": float(row["margin"]),
            "margin_pct": safe_divide(float(row["margin"]), float(row["net_sales"])),
            "units": float(row["units"]),
            "avg_sales_price": safe_divide(float(row["net_sales"]), float(row["units"])),
        })
    return pd.DataFrame(rows).sort_values("group", key=lambda s: s.map(DEPARTMENT_ORDER).fillna(99))


def build_sub_department_summary(executive_kpis: pd.DataFrame, department_sales: pd.DataFrame, department_targets: pd.DataFrame) -> pd.DataFrame:
    total_sales = float(executive_kpis["net_sales"].iloc[0]) if not executive_kpis.empty else department_sales["net_sales"].sum()
    detail = pd.merge(department_sales, department_targets, on=["group", "department"], how="outer")
    detail["net_sales"] = detail["net_sales"].fillna(0.0)
    detail["budget"] = detail["budget"].fillna(0.0)
    detail["units"] = detail["units"].fillna(0.0)
    detail["budget_gap"] = detail["net_sales"] - detail["budget"]
    detail["budget_gap_pct"] = [safe_divide(g, b) for g, b in zip(detail["budget_gap"], detail["budget"])]
    detail["mix_pct"] = [safe_divide(s, total_sales) for s in detail["net_sales"]]
    detail["avg_sales_price"] = [safe_divide(s, u) for s, u in zip(detail["net_sales"], detail["units"])]
    detail["priority_score"] = [abs(gp) * m if gp is not None else 0 for gp, m in zip(detail["budget_gap_pct"], detail["mix_pct"])]
    detail["sort_order"] = detail["department"].map(SUBDEPT_SORT_ORDER).fillna(99)
    return detail[[
        "group", "department", "net_sales", "budget", "budget_gap", "budget_gap_pct",
        "mix_pct", "net_sales_mix_within_group_pct", "units", "avg_sales_price", "priority_score", "sort_order"
    ]].sort_values(["sort_order", "group", "department"]).drop(columns=["sort_order"])


def build_all_brands(brand_detail: pd.DataFrame) -> pd.DataFrame:
    df = brand_detail.copy()
    df["rank"] = df.groupby(["group", "department"])["net_sales"].rank(ascending=False, method="first").astype(int)
    df = df.sort_values(["group", "department", "rank"])
    return df[["group", "department", "rank", "brand", "net_sales", "units", "department_brand_mix_pct", "margin", "margin_pct"]]


def build_top_brands_fixed(all_brands: pd.DataFrame, top_n: int = 5) -> pd.DataFrame:
    rows = []
    for group, sub in SUBDEPARTMENT_ORDER:
        sub_df = all_brands[(all_brands["group"] == group) & (all_brands["department"] == sub)].sort_values("rank")
        for rank in range(1, top_n + 1):
            match = sub_df[sub_df["rank"] == rank]
            if not match.empty:
                r = match.iloc[0]
                rows.append({"group": group, "department": sub, "rank": rank, "brand": r["brand"], "net_sales": r["net_sales"], "mix_pct": r["department_brand_mix_pct"], "units": r["units"]})
            else:
                rows.append({"group": group, "department": sub, "rank": rank, "brand": "", "net_sales": 0.0, "mix_pct": 0.0, "units": 0.0})
    return pd.DataFrame(rows)


def build_chart_data(department_summary: pd.DataFrame, sub_department_summary: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    return {
        "chart_sales_mix": department_summary[["group", "net_sales_mix_pct"]].rename(columns={"group": "department", "net_sales_mix_pct": "mix_pct"}),
        "chart_budget_gap_department": department_summary[["group", "budget_gap"]].rename(columns={"group": "department"}),
        "chart_budget_gap_subdepartment": sub_department_summary[["department", "budget_gap", "budget_gap_pct"]].copy(),
    }


def build_kpi_master(executive_kpis: pd.DataFrame, department_summary: pd.DataFrame, traffic_summary: pd.DataFrame) -> pd.DataFrame:
    e = executive_kpis.iloc[0]
    sales = float(e["net_sales"])
    budget = float(department_summary["budget"].sum())
    budget_gap = sales - budget
    footfall = float(traffic_summary.loc[traffic_summary["metric"] == "Total Footfall", "value"].iloc[0]) if not traffic_summary.empty else 0.0
    traffic_transactions = float(traffic_summary.loc[traffic_summary["metric"] == "Total Transactions", "value"].iloc[0]) if not traffic_summary.empty else 0.0

    # Safety fallback: if the hourly traffic parser ever fails, use the already validated LFL totals.
    # This prevents executive KPIs from showing zero when LFL contains the correct values.
    if footfall == 0 and "footfall" in e:
        footfall = float(e.get("footfall", 0) or 0)
    if traffic_transactions == 0:
        traffic_transactions = float(e.get("transactions_lfl_source", 0) or 0)

    conversion_pct = safe_divide(traffic_transactions, footfall)
    units = float(department_summary["units"].sum())
    margin = float(department_summary["margin"].sum())

    rows = [
        ["Net Sales", sales, float(e["py_net_sales"]), float(e["net_sales_py_var"]), float(e["sales_lfl_pct"]), "LFL.csv"],
        ["Budget", budget, None, None, None, "Department summary total budget"],
        ["Budget Gap", budget_gap, None, budget_gap, safe_divide(budget_gap, budget), "Net Sales - Budget"],
        ["Budget Gap %", safe_divide(budget_gap, budget), None, None, safe_divide(budget_gap, budget), "Budget Gap / Budget"],
        ["Budget Attainment %", safe_divide(sales, budget), None, None, None, "Net Sales / Budget"],
        ["Sales vs LY", float(e["net_sales_py_var"]), float(e["py_net_sales"]), float(e["net_sales_py_var"]), float(e["sales_lfl_pct"]), "LFL.csv"],
        ["Footfall", footfall, None, None, None, "Table.csv"],
        ["Transactions", traffic_transactions, float(e["py_transactions"]), float(e["transactions_py_var"]), float(e["transactions_lfl_pct"]), "Table.csv / LFL.csv"],
        ["Conversion %", conversion_pct, None, None, None, "Transactions / Footfall from Table.csv"],
        ["IPC / UPT", float(e["ipc"]), float(e["py_ipc"]), float(e["ipc_py_var"]), float(e["ipc_lfl_pct"]), "LFL.csv / Tableau source"],
        ["Units", units, None, None, None, "Group by Site.csv"],
        ["Avg Price / AUR", safe_divide(sales, units), None, None, None, "Net Sales / Units"],
        ["ATV", safe_divide(sales, traffic_transactions), None, None, None, "Net Sales / Transactions"],
        ["Margin $", margin, None, None, None, "Group by Site.csv"],
        ["Margin %", safe_divide(margin, sales), None, None, None, "Margin $ / Net Sales"],
    ]
    return pd.DataFrame(rows, columns=["kpi", "value", "ly_or_target", "variance", "variance_pct", "source_notes"])


def build_report_facts(kpi_master: pd.DataFrame, department_summary: pd.DataFrame, sub_department_summary: pd.DataFrame, all_brands: pd.DataFrame, traffic_summary: pd.DataFrame) -> pd.DataFrame:
    largest_dept = department_summary.sort_values("budget_gap").iloc[0]
    largest_sub = sub_department_summary.sort_values("budget_gap").iloc[0]
    best_mix_dept = department_summary.sort_values("net_sales_mix_pct", ascending=False).iloc[0]
    best_sub = sub_department_summary.sort_values("budget_gap_pct", ascending=False).iloc[0]
    worst_sub = sub_department_summary.sort_values("budget_gap_pct", ascending=True).iloc[0]
    top_brand = all_brands.sort_values("net_sales", ascending=False).iloc[0] if not all_brands.empty else None

    def traffic_value(metric: str) -> str:
        row = traffic_summary[traffic_summary["metric"] == metric]
        return "" if row.empty else str(row.iloc[0]["value"])

    key_takeaway = (
        f"Focus on {largest_sub['department']} and broader {largest_dept['group']} execution while protecting stronger categories and improving conversion."
    )

    rows = [
        ["Largest Budget Gap Department", largest_dept["group"], "DEPARTMENT_SUMMARY", f"{largest_dept['budget_gap']:,.0f}"],
        ["Largest Budget Gap SubDepartment", largest_sub["department"], "SUB_DEPARTMENT_SUMMARY", f"{largest_sub['budget_gap']:,.0f}"],
        ["Best Sales Mix Department", best_mix_dept["group"], "DEPARTMENT_SUMMARY", f"{best_mix_dept['net_sales_mix_pct']:.1%}"],
        ["Best SubDepartment vs Budget", best_sub["department"], "SUB_DEPARTMENT_SUMMARY", f"{best_sub['budget_gap_pct']:.1%}"],
        ["Worst SubDepartment vs Budget", worst_sub["department"], "SUB_DEPARTMENT_SUMMARY", f"{worst_sub['budget_gap_pct']:.1%}"],
        ["Top Brand Overall", top_brand["brand"] if top_brand is not None else "", "ALL_BRANDS_DETAIL", f"{top_brand['department']} | ${top_brand['net_sales']:,.0f}" if top_brand is not None else ""],
        ["Peak Traffic Hour", traffic_value("Peak Traffic Hour"), "HOURLY_TRAFFIC", ""],
        ["Best Conversion Hour", traffic_value("Best Conversion Hour"), "HOURLY_TRAFFIC", ""],
        ["Key Takeaway", key_takeaway, "Generated", ""],
    ]
    return pd.DataFrame(rows, columns=["metric", "value", "source", "notes"])


def build_validation(executive_kpis: pd.DataFrame, department_summary: pd.DataFrame, sub_department_summary: pd.DataFrame, all_brands: pd.DataFrame, hourly_traffic: pd.DataFrame) -> pd.DataFrame:
    sales = float(executive_kpis["net_sales"].iloc[0])
    dept_sales = float(department_summary["net_sales"].sum())
    sub_sales = float(sub_department_summary["net_sales"].sum())
    brand_sales = float(all_brands["net_sales"].sum())
    traffic_sales = float(hourly_traffic["net_sales"].sum()) if not hourly_traffic.empty else 0.0
    rows = [
        ["Executive vs Department Sales", sales, dept_sales, sales - dept_sales, abs(sales - dept_sales) <= 10],
        ["Department vs SubDepartment Sales", dept_sales, sub_sales, dept_sales - sub_sales, abs(dept_sales - sub_sales) <= 10],
        ["Department vs Brand Sales", dept_sales, brand_sales, dept_sales - brand_sales, abs(dept_sales - brand_sales) <= 10],
        ["Executive vs Traffic Hourly Sales", sales, traffic_sales, sales - traffic_sales, abs(sales - traffic_sales) <= 2000],
    ]
    return pd.DataFrame(rows, columns=["check", "source_a", "source_b", "difference", "passed"])


def build_model(base_dir: str | Path = ".") -> Dict[str, pd.DataFrame]:
    base = Path(base_dir)
    raw = {name: read_tableau_csv(base / fname) for name, fname in INPUT_FILES.items()}

    executive_kpis = clean_lfl(raw["lfl"])
    dept_summary_raw = clean_department_summary(raw["group_by_site"])
    department_targets, dept_target_totals = clean_department_targets(raw["d_net_sales"])
    department_sales = clean_department_sales(raw["d_net_sales"])
    hourly_traffic, traffic_summary = clean_traffic_by_hour(raw["traffic_by_hour"])
    option_list_detail = clean_option_list(raw["option_list"])
    top_sellers_global, top_sellers_department, top_sellers_subdepartment = build_top_sellers(option_list_detail)

    department_summary = build_department_summary(executive_kpis, dept_summary_raw, dept_target_totals)
    sub_department_summary = build_sub_department_summary(executive_kpis, department_sales, department_targets)

    brand_detail = pd.concat([
        clean_brand_file(raw["footwear"], "Footwear"),
        clean_brand_file(raw["apparel"], "Apparel"),
        clean_brand_file(raw["accessories"], "Accessories"),
    ], ignore_index=True)
    all_brands_detail = build_all_brands(brand_detail)
    top_brands_fixed = build_top_brands_fixed(all_brands_detail)
    chart_tables = build_chart_data(department_summary, sub_department_summary)
    kpi_master = build_kpi_master(executive_kpis, department_summary, traffic_summary)
    report_facts = build_report_facts(kpi_master, department_summary, sub_department_summary, all_brands_detail, traffic_summary)
    validation = build_validation(executive_kpis, department_summary, sub_department_summary, all_brands_detail, hourly_traffic)

    # Placeholder until the district PDF parser is added.
    district_benchmark = pd.DataFrame({
        "category": ["Mens Apparel", "Womens Apparel", "Junior Apparel", "Mens Footwear", "Womens Footwear", "Junior Footwear", "Clothing Accessories", "Other Accessories"],
        "store_pct": [0.0] * 8,
        "district_pct": [0.0] * 8,
        "difference": [0.0] * 8,
        "status": ["District PDF pending"] * 8,
        "source": ["District PDF"] * 8,
    })

    report_info = pd.DataFrame([
        ["Store Name", "Richmond Centre"],
        ["Report Period", "Current selected period"],
        ["Data Model Version", "1.0"],
    ], columns=["field", "value"])

    return {
        "report_info": report_info,
        "executive_kpis": executive_kpis,
        "kpi_master": kpi_master,
        "department_summary": department_summary,
        "sub_department_summary": sub_department_summary,
        **chart_tables,
        "top_brands_fixed": top_brands_fixed,
        "all_brands_detail": all_brands_detail,
        "district_benchmark": district_benchmark,
        "traffic_summary": traffic_summary,
        "hourly_traffic": hourly_traffic,
        "report_facts": report_facts,
        "validation": validation,
        "option_list_detail": option_list_detail,
        "top_sellers_global": top_sellers_global,
        "top_sellers_department": top_sellers_department,
        "top_sellers_subdepartment": top_sellers_subdepartment,
    }


def export_model(tables: Dict[str, pd.DataFrame], output_path: str | Path) -> None:
    output_path = Path(output_path)
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})
        money_fmt = workbook.add_format({"num_format": "$#,##0;[Red]-$#,##0"})
        pct_fmt = workbook.add_format({"num_format": "0.0%;[Red]-0.0%"})
        int_fmt = workbook.add_format({"num_format": "#,##0"})

        for sheet_name, df in tables.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
            ws = writer.sheets[sheet_name[:31]]
            ws.freeze_panes(1, 0)
            for idx, col in enumerate(df.columns):
                ws.write(0, idx, col, header_fmt)
                col_l = str(col).lower()
                width = min(max(len(str(col)) + 4, 14), 32)
                fmt = None
                if any(k in col_l for k in ["sales", "budget", "gap", "margin", "price", "value", "variance"]):
                    fmt = money_fmt
                if "pct" in col_l or "conversion" in col_l or "mix" in col_l:
                    fmt = pct_fmt
                if any(k in col_l for k in ["units", "transactions", "footfall", "rank"]):
                    fmt = int_fmt
                ws.set_column(idx, idx, width, fmt)


if __name__ == "__main__":
    tables = build_model("Sample Inputs")
    export_model(tables, "vm_weekly_report_data_model.xlsx")
    print("Created vm_weekly_report_data_model.xlsx")
