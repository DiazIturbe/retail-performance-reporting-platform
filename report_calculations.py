"""
report_calculations.py
Weekly VM Operational Report - Calculations and summary engine.

This version consumes the updated data model produced by vm_report_data_model_v2.py,
including Table.csv / Traffic By The Hour data.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict

import pandas as pd


def safe_divide(numerator: float, denominator: float, default: float = 0.0) -> float:
    try:
        if denominator in (0, None) or pd.isna(denominator):
            return default
        return numerator / denominator
    except Exception:
        return default


def money(value: float) -> str:
    if value is None or pd.isna(value):
        return "$0"
    sign = "-" if value < 0 else ""
    return f"{sign}${abs(value):,.0f}"


def pct(value: float, decimals: int = 1) -> str:
    if value is None or pd.isna(value):
        return "0.0%"
    return f"{value * 100:.{decimals}f}%"


def number(value: float, decimals: int = 0) -> str:
    if value is None or pd.isna(value):
        value = 0
    return f"{value:,.{decimals}f}"


def load_data_model(input_path: str | Path) -> Dict[str, pd.DataFrame]:
    input_path = Path(input_path)
    workbook = pd.ExcelFile(input_path)
    required = [
        "kpi_master",
        "department_summary",
        "sub_department_summary",
        "top_brands_fixed",
        "all_brands_detail",
        "district_benchmark",
        "traffic_summary",
        "hourly_traffic",
        "report_facts",
        "validation",
    ]
    missing = [s for s in required if s not in workbook.sheet_names]
    if missing:
        raise ValueError(f"Missing required data-model sheets: {missing}")
    data = {sheet: pd.read_excel(input_path, sheet_name=sheet) for sheet in required}
    for sheet in ["option_list_detail", "top_sellers_global", "top_sellers_department", "top_sellers_subdepartment"]:
        data[sheet] = pd.read_excel(input_path, sheet_name=sheet) if sheet in workbook.sheet_names else pd.DataFrame()
    return data


def _metric_value(kpi_master: pd.DataFrame, name: str, default=0.0):
    row = kpi_master[kpi_master["kpi"].astype(str).str.lower() == name.lower()]
    if row.empty:
        return default
    return row.iloc[0].get("value", default)


def build_opportunity_ranking(sub_department_summary: pd.DataFrame, top_n: int = 8) -> pd.DataFrame:
    df = sub_department_summary.copy()
    df = df[df["budget_gap"] < 0].copy()
    df = df.sort_values("priority_score", ascending=False).head(top_n)
    df["recommendation"] = df.apply(make_recommendation, axis=1)
    df["priority_score_note"] = "Gap severity weighted by sales mix"
    return df[["group", "department", "net_sales", "budget", "budget_gap", "budget_gap_pct", "mix_pct", "priority_score", "priority_score_note", "recommendation"]]


def make_recommendation(row: pd.Series) -> str:
    group = str(row.get("group", "")).lower()
    dept = str(row.get("department", ""))
    if "footwear" in group:
        return f"Prioritize size availability, wall standards, replenishment, and add-on selling for {dept}."
    if "apparel" in group:
        return f"Review outfit building, fixture density, size runs, and front-of-store visibility for {dept}."
    if "accessor" in group:
        return f"Increase cash-wrap visibility, add-on scripting, and replenishment frequency for {dept}."
    return f"Review availability, placement, and selling focus for {dept}."


def generate_executive_summary(data: Dict[str, pd.DataFrame], store_name: str = "Store") -> str:
    kpi = data["kpi_master"]
    dept = data["department_summary"]
    sub = data["sub_department_summary"]
    facts = data["report_facts"]

    net_sales = _metric_value(kpi, "Net Sales")
    budget_gap_pct = _metric_value(kpi, "Budget Gap %")
    sales_vs_ly_pct = kpi.loc[kpi["kpi"] == "Sales vs LY", "variance_pct"].iloc[0] if not kpi.loc[kpi["kpi"] == "Sales vs LY"].empty else 0
    conversion = _metric_value(kpi, "Conversion %")
    footfall = _metric_value(kpi, "Footfall")
    atv = _metric_value(kpi, "ATV")
    ipc = _metric_value(kpi, "IPC / UPT")

    top_dept = dept.sort_values("net_sales_mix_pct", ascending=False).iloc[0]
    largest_gap = sub.sort_values("budget_gap").iloc[0]
    best_sub = sub.sort_values("budget_gap_pct", ascending=False).iloc[0]

    budget_phrase = "above budget" if budget_gap_pct >= 0 else "below budget"
    ly_phrase = "above last year" if sales_vs_ly_pct >= 0 else "below last year"

    return (
        f"{store_name} generated {money(net_sales)} in net sales, finishing "
        f"{pct(abs(budget_gap_pct))} {budget_phrase} and {pct(abs(sales_vs_ly_pct))} {ly_phrase}. "
        f"{top_dept['group']} represented {pct(top_dept['net_sales_mix_pct'])} of total sales and remains the primary sales driver. "
        f"The largest budget opportunity is {largest_gap['department']} with a gap of {money(largest_gap['budget_gap'])}, "
        f"while {best_sub['department']} is the strongest subdepartment versus budget. "
        f"Footfall was {number(footfall)} with conversion at {pct(conversion)}, ATV at {money(atv)}, and IPC at {ipc:.2f}."
    )


def generate_department_insights(department_summary: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in department_summary.iterrows():
        direction = "above budget" if row["budget_gap"] >= 0 else "below budget"
        rows.append({
            "group": row["group"],
            "insight": (
                f"{row['group']} generated {money(row['net_sales'])}, representing {pct(row['net_sales_mix_pct'])} of total sales. "
                f"It finished {money(abs(row['budget_gap']))} {direction} ({pct(row['budget_gap_pct'])})."
            )
        })
    return pd.DataFrame(rows)



def build_top_brands_overall(all_brands_detail: pd.DataFrame, top_n: int = 15) -> pd.DataFrame:
    """Aggregate brand sales by department category and brand before ranking.

    The raw Tableau brand exports contain one row per subdepartment-brand pair.
    Sorting those rows directly makes Nike/Jordan/adidas look artificially small
    because their sales are split across Mens/Womens/Junior/Kids.
    This function first groups by group + brand, then ranks the total.
    """
    df = all_brands_detail.copy()
    if df.empty:
        return df

    grouped = (
        df.groupby(["group", "brand"], as_index=False)
        .agg({
            "net_sales": "sum",
            "units": "sum",
            "margin": "sum",
        })
    )
    grouped["margin_pct"] = grouped.apply(
        lambda r: safe_divide(float(r["margin"]), float(r["net_sales"])), axis=1
    )

    # Mix within the full report total, useful for the top-brand chart/table.
    total_sales = float(grouped["net_sales"].sum())
    grouped["mix_pct"] = grouped["net_sales"].apply(lambda x: safe_divide(float(x), total_sales))

    # Rank brand/category combinations. A brand can appear more than once
    # when it performs in more than one department category (Nike Footwear, Nike Apparel).
    grouped = grouped.sort_values("net_sales", ascending=False).head(top_n).copy()
    grouped.insert(0, "rank", range(1, len(grouped) + 1))
    return grouped[["rank", "group", "brand", "net_sales", "mix_pct", "units", "margin", "margin_pct"]]

DEPARTMENT_ORDER = {"footwear": 0, "apparel": 1, "accessories": 2}
SUBDEPARTMENT_ORDER = {
    "mens": 0, "men's": 0, "womens": 1, "women's": 1,
    "junior": 2, "childrens": 3, "children's": 3,
    "infants": 4, "infant": 4, "clothing accessories": 5, "other accessories": 6,
}


def _department_order(value):
    return DEPARTMENT_ORDER.get(str(value).strip().lower(), 99)


def _subdepartment_order(value):
    text = str(value).strip().lower()
    for key, rank in SUBDEPARTMENT_ORDER.items():
        if key in text:
            return rank
    return 99


def _sort_report_frames(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    for name in ["department_summary", "top_sellers_department"]:
        df = data.get(name)
        if isinstance(df, pd.DataFrame) and not df.empty and "group" in df.columns:
            df = df.copy()
            df["_group_order"] = df["group"].apply(_department_order)
            sort_cols = ["_group_order"]
            if "rank" in df.columns:
                sort_cols.append("rank")
            data[name] = df.sort_values(sort_cols).drop(columns="_group_order")

    for name in ["sub_department_summary", "top_brands_fixed", "top_sellers_subdepartment"]:
        df = data.get(name)
        if isinstance(df, pd.DataFrame) and not df.empty and "group" in df.columns:
            df = df.copy()
            df["_group_order"] = df["group"].apply(_department_order)
            if "department" in df.columns:
                df["_sub_order"] = df["department"].apply(_subdepartment_order)
            else:
                df["_sub_order"] = 99
            sort_cols = ["_group_order", "_sub_order"]
            if "rank" in df.columns:
                sort_cols.append("rank")
            data[name] = df.sort_values(sort_cols).drop(columns=["_group_order", "_sub_order"])
    return data


def build_report_calculations(input_path: str | Path, store_name: str = "Store") -> Dict[str, pd.DataFrame | str]:
    data = _sort_report_frames(load_data_model(input_path))
    department_style_map = pd.DataFrame([
        {"group": "Footwear", "order": 1, "color_key": "footwear"},
        {"group": "Apparel", "order": 2, "color_key": "apparel"},
        {"group": "Accessories", "order": 3, "color_key": "accessories"},
    ])

    outputs: Dict[str, pd.DataFrame | str] = {
        "kpi_master": data["kpi_master"],
        "department_kpis": data["department_summary"],
        "sub_department_kpis": data["sub_department_summary"],
        "opportunity_ranking": build_opportunity_ranking(data["sub_department_summary"]),
        "top_brands_fixed": data["top_brands_fixed"],
        "top_brands_overall": build_top_brands_overall(data["all_brands_detail"], top_n=15),
        "traffic_summary": data["traffic_summary"],
        "hourly_traffic": data["hourly_traffic"],
        "department_insights": generate_department_insights(data["department_summary"]),
        "executive_summary_text": generate_executive_summary(data, store_name=store_name),
        "validation": data["validation"],
        "top_sellers_global": data.get("top_sellers_global", pd.DataFrame()),
        "top_sellers_department": data.get("top_sellers_department", pd.DataFrame()),
        "top_sellers_subdepartment": data.get("top_sellers_subdepartment", pd.DataFrame()),
        "department_style_map": department_style_map,
    }
    return outputs


def export_calculations_to_excel(outputs: Dict[str, pd.DataFrame | str], output_path: str | Path) -> None:
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})
        money_fmt = workbook.add_format({"num_format": "$#,##0;[Red]-$#,##0"})
        pct_fmt = workbook.add_format({"num_format": "0.0%;[Red]-0.0%"})
        int_fmt = workbook.add_format({"num_format": "#,##0"})

        for sheet_name, value in outputs.items():
            clean = sheet_name[:31]
            if isinstance(value, pd.DataFrame):
                df = value
            else:
                df = pd.DataFrame({"text": [value]})
            df.to_excel(writer, index=False, sheet_name=clean)
            ws = writer.sheets[clean]
            ws.freeze_panes(1, 0)
            for idx, col in enumerate(df.columns):
                ws.write(0, idx, col, header_fmt)
                col_l = str(col).lower()
                fmt = None
                if any(k in col_l for k in ["sales", "budget", "gap", "margin", "price", "value", "variance"]):
                    fmt = money_fmt
                if "pct" in col_l or "conversion" in col_l or "mix" in col_l:
                    fmt = pct_fmt
                if any(k in col_l for k in ["units", "transactions", "footfall", "rank"]):
                    fmt = int_fmt
                ws.set_column(idx, idx, 18, fmt)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build weekly VM report calculations.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=False)
    args = parser.parse_args()
    outputs = build_report_calculations(args.input)
    print(outputs["executive_summary_text"])
    if args.output:
        export_calculations_to_excel(outputs, args.output)
        print(f"Saved {args.output}")


if __name__ == "__main__":
    main()
