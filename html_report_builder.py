from pathlib import Path
import re

from report_calculations import build_report_calculations
from chart_builder import (
    make_budget_gap_chart,
    make_sales_mix_donut,
    make_top_brands_chart,
    make_subdepartment_mix_chart,
)

OUTPUT_DIR = Path("report_output")
OUTPUT_DIR.mkdir(exist_ok=True)

DATA_MODEL = "vm_weekly_report_data_model.xlsx"


def money(value):
    try:
        return f"${float(value):,.0f}"
    except Exception:
        return "$0"


def money_signed(value):
    try:
        value = float(value)
        sign = "-" if value < 0 else ""
        return f"{sign}${abs(value):,.0f}"
    except Exception:
        return "$0"


def pct(value):
    try:
        return f"{float(value) * 100:.1f}%"
    except Exception:
        return "0.0%"


def number(value):
    try:
        return f"{float(value):,.0f}"
    except Exception:
        return "0"


def get_kpi_value(kpi, name, column="value", default=0):
    row = kpi[kpi["kpi"].astype(str).str.lower() == name.lower()]
    if row.empty or column not in row.columns:
        return default
    return row.iloc[0].get(column, default)


def build_department_table(dept):
    display = dept.copy()
    if display.empty:
        return '<div class="empty-note">No department data available</div>'

    display["_order"] = display["group"].apply(department_sort_key)
    display = display.sort_values("_order").drop(columns="_order")

    max_gap = max(display["budget_gap"].abs().max(), 1)
    max_mix = max(display["net_sales_mix_pct"].abs().max(), 0.000001)
    rows = []
    for _, row in display.iterrows():
        css = department_class(row.get("group", ""))
        rows.append(f"""
        <tr class="department-row {css}-row">
            <td class="text-left department-name-cell"><span class="department-pill {css}">{section_icon(row.get('group', ''), 'table-dept-icon')}{row.get('group', '')}</span></td>
            <td>{money(row.get('net_sales', 0))}</td>
            <td>{money(row.get('budget', 0))}</td>
            <td>{format_spark_cell(row.get('budget_gap', 0), money_signed(row.get('budget_gap', 0)), kind='gap', max_abs=max_gap)}</td>
            <td>{format_colored_value(row.get('budget_gap_pct', 0), pct(row.get('budget_gap_pct', 0)))}</td>
            <td>{format_spark_cell(row.get('net_sales_mix_pct', 0), pct(row.get('net_sales_mix_pct', 0)), kind=css, max_abs=max_mix)}</td>
            <td>{number(row.get('units', 0))}</td>
            <td>{money(row.get('avg_sales_price', 0))}</td>
        </tr>
        """)

    return f"""
    <table class="data-table department-performance-table visual-table">
        <thead><tr>
            <th>Department</th><th>Net Sales</th><th>Budget</th><th>Budget Gap</th>
            <th>Gap %</th><th>Contribution to Total Sales</th><th>Units</th><th>ATV</th>
        </tr></thead>
        <tbody>{''.join(rows)}</tbody>
    </table>
    """


def build_top_brands_table(brands, top_n=8):
    """Top brand table with a clear brand-category label.

    A brand can be important in more than one department, e.g. Nike Footwear
    and Nike Apparel. To avoid the appearance of duplicate brands, this table
    combines Brand + Department into one label.
    """
    display = brands.copy().head(top_n)
    if display.empty:
        return '<div class="empty-note">No brand data available</div>'

    group_col = "group" if "group" in display.columns else ("Department" if "Department" in display.columns else None)
    brand_col = "brand" if "brand" in display.columns else "Brand"
    sales_col = "net_sales" if "net_sales" in display.columns else "Sales"
    mix_col = "mix_pct" if "mix_pct" in display.columns else None
    units_col = "units" if "units" in display.columns else None

    display[brand_col] = display[brand_col].astype(str).str.strip()
    if group_col:
        display[group_col] = display[group_col].astype(str).str.strip()
        display["Brand Category"] = display[brand_col] + " — " + display[group_col]
    else:
        display["Brand Category"] = display[brand_col]

    cols = ["Brand Category", sales_col]
    if mix_col:
        cols.append(mix_col)
    if units_col:
        cols.append(units_col)

    display = display[cols].copy()
    rename = {sales_col: "Net Sales"}
    if mix_col:
        rename[mix_col] = "Contribution"
    if units_col:
        rename[units_col] = "Units"
    display = display.rename(columns=rename)

    display["Net Sales"] = display["Net Sales"].apply(money)
    if "Contribution" in display.columns:
        display["Contribution"] = display["Contribution"].apply(pct)
    if "Units" in display.columns:
        display["Units"] = display["Units"].apply(number)

    return display.to_html(index=False, classes="data-table brand-performance-table", border=0)


def build_opportunities_table(opportunities):
    display = opportunities.copy()
    rename = {
        "group": "Group",
        "department": "Department",
        "net_sales": "Net Sales",
        "budget": "Budget",
        "budget_gap": "Budget Gap",
        "budget_gap_pct": "Gap %",
        "mix_pct": "Mix %",
        "priority_score": "Priority Score",
        "recommendation": "Recommendation",
    }
    display = display.rename(columns={k: v for k, v in rename.items() if k in display.columns})
    for col in ["Net Sales", "Budget"]:
        if col in display.columns:
            display[col] = display[col].apply(money)
    if "Budget Gap" in display.columns:
        display["Budget Gap"] = display["Budget Gap"].apply(money_signed)
    for col in ["Gap %", "Mix %"]:
        if col in display.columns:
            display[col] = display[col].apply(pct)
    if "Priority Score" in display.columns:
        display["Priority Score"] = display["Priority Score"].apply(lambda x: f"{float(x):.2f}" if str(x) != "nan" else "0.00")
    return display.to_html(index=False, classes="data-table opportunities-table", border=0)


def build_validation_table(validation):
    display = validation.copy()
    return display.to_html(index=False, classes="data-table", border=0)


def department_class(name):
    name = str(name).lower()
    if "footwear" in name:
        return "footwear"
    if "apparel" in name:
        return "apparel"
    if "access" in name:
        return "accessories"
    return "neutral"


DEPARTMENT_ORDER = {"footwear": 0, "apparel": 1, "accessories": 2}
SUBDEPARTMENT_ORDER = {
    "mens": 0, "men's": 0,
    "womens": 1, "women's": 1,
    "junior": 2,
    "childrens": 3, "children's": 3,
    "infants": 4, "infant": 4,
    "clothing accessories": 5,
    "other accessories": 6,
}


def department_sort_key(value):
    return DEPARTMENT_ORDER.get(department_class(value), 99)


def subdepartment_sort_key(value):
    text = str(value).strip().lower()
    for key, rank in SUBDEPARTMENT_ORDER.items():
        if key in text:
            return rank
    return 99


def section_icon(name, css_class="section-icon"):
    icon_name = department_class(name)
    paths = {
        "footwear": '<path d="M3 15c4.5.2 7.2-1.2 9.5-4.3l2.2 1.1c1.2.6 2.7 1 4.3 1.2l1.8.2c.7.1 1.2.7 1.2 1.4v1.2c0 1.2-1 2.2-2.2 2.2H6.2A3.2 3.2 0 0 1 3 14.8z"/><path d="M12.5 10.7 10 7.5m5.2 4.4-1.8-3.6"/>',
        "apparel": '<path d="M8 4 4 6.5l2.2 4L8 9.5V20h8V9.5l1.8 1 2.2-4L16 4l-2 2h-4z"/>',
        "accessories": '<path d="M6 8h12l1.5 12h-15z"/><path d="M9 8V6a3 3 0 0 1 6 0v2"/>',
        "neutral": '<circle cx="12" cy="12" r="8"/><path d="M8 12h8m-4-4v8"/>',
    }
    return f'<svg class="{css_class}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">{paths[icon_name]}</svg>'


def subsection_icon(name):
    text = str(name).lower()
    if "women" in text:
        path = '<circle cx="12" cy="6" r="2.4"/><path d="M8.2 19.5 10 11h4l1.8 8.5M9.2 16h5.6"/>'
    elif "men" in text:
        path = '<circle cx="12" cy="6" r="2.5"/><path d="M7.5 20v-4.8a4.5 4.5 0 0 1 9 0V20M9.5 12.2 7 15m7.5-2.8L17 15"/>'
    elif "junior" in text:
        path = '<circle cx="12" cy="7" r="2.1"/><path d="M8.7 19.5v-4.2a3.3 3.3 0 0 1 6.6 0v4.2m-5-7.1-2.3 2.1m5.7-2.1 2.3 2.1"/>'
    elif "child" in text:
        path = '<circle cx="12" cy="7.5" r="1.9"/><path d="M9 19v-4a3 3 0 0 1 6 0v4m-4.4-6.2-2 2.6m4.8-2.6 2 2.6"/>'
    elif "infant" in text or "baby" in text:
        path = '<path d="M7 11.5h10a5 5 0 0 1-10 0Z"/><path d="M12 6a4 4 0 0 1 4 4H8a4 4 0 0 1 4-4Zm-3.5 12.5h7M10 20h4"/>'
    elif "clothing" in text:
        path = '<path d="M9 4h6v5l2 3-2 8H9l-2-8 2-3Z"/><path d="M9 8h6"/>'
    else:
        path = '<path d="M6 8h12l1.5 12h-15z"/><path d="M9 8V6a3 3 0 0 1 6 0v2"/>'
    return f'<svg class="subsection-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">{path}</svg>'


def highlight_summary(text):
    """Add restrained visual emphasis without changing the generated wording."""
    safe = str(text)
    safe = re.sub(r'(\$[\d,]+(?:\.\d+)?)', r'<span class="summary-chip money-chip">\1</span>', safe)
    safe = re.sub(
        r'(\d+(?:\.\d+)?%\s+(?:above|below)\s+(?:budget|last year))',
        lambda m: f'<span class="summary-chip {"positive-chip" if "above" in m.group(1) else "negative-chip"}">{m.group(1)}</span>',
        safe,
        flags=re.IGNORECASE,
    )
    safe = re.sub(r'(\d+(?:\.\d+)?%\s+of total sales)', r'<span class="summary-chip mix-chip">\1</span>', safe, flags=re.IGNORECASE)
    return safe


def build_brand_column(brands, department_name, group_class='neutral', shade_index=0):
    filtered = brands[
        brands["department"].astype(str).str.lower() == department_name.lower()
    ].copy()
    filtered = filtered.dropna(subset=["brand"])
    filtered = filtered[filtered["net_sales"] > 0]
    filtered = filtered.sort_values("rank").head(5)

    if filtered.empty:
        return f"""
        <div class="brand-column {group_class}-column">
            <div class="brand-column-title">{subsection_icon(department_name)}<span>{department_name}</span></div>
            <div class="empty-note">No brand data available</div>
        </div>
        """

    rows = ""
    for _, row in filtered.iterrows():
        rows += f"""
        <tr>
            <td>{row['brand']}</td>
            <td>{money(row['net_sales'])}</td>
            <td>{number(row['units'])}</td>
        </tr>
        """

    shade_class = f"{group_class}-shade-{min(shade_index + 1, 5)}"
    return f"""
    <div class="brand-column {group_class}-column {shade_class}">
        <div class="brand-column-title">{subsection_icon(department_name)}<span>{department_name}</span></div>
        <table class="mini-table">
            <thead>
                <tr><th>Brand</th><th>Sales</th><th>Units</th></tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>
    </div>
    """


def build_department_card(row, top_brands_fixed, top_sellers_department=None):
    dept_name = row["group"]
    css_class = department_class(dept_name)

    department_names = (
        top_brands_fixed[
            top_brands_fixed["group"].astype(str).str.lower() == str(dept_name).lower()
        ]["department"]
        .dropna()
        .unique()
        .tolist()
    )
    department_names = sorted(department_names, key=subdepartment_sort_key)

    brand_columns = "\n".join(
        build_brand_column(top_brands_fixed, department_name, css_class, idx)
        for idx, department_name in enumerate(department_names)
    )

    seller_html = ""
    if top_sellers_department is not None and not top_sellers_department.empty:
        seller_rows = top_sellers_department[
            top_sellers_department["group"].astype(str).str.lower() == str(dept_name).lower()
        ]
        if not seller_rows.empty:
            seller_html = f"""
            <div class="embedded-sellers {css_class}-seller">
                <div class="embedded-sellers-title">{section_icon(dept_name, 'seller-title-icon')}<span>Top Sellers — {dept_name}</span></div>
                {build_top_sellers_table(seller_rows, top_n=3)}
            </div>
            """

    return f"""
    <div class="department-card {css_class}-department-card">
        <div class="department-header {css_class}">{section_icon(dept_name)}<span>{dept_name}</span></div>

        <div class="metric-row">
            <div class="small-metric"><div class="small-metric-label">Net Sales</div><div class="small-metric-value">{money(row['net_sales'])}</div></div>
            <div class="small-metric"><div class="small-metric-label">Budget</div><div class="small-metric-value">{money(row['budget'])}</div></div>
            <div class="small-metric"><div class="small-metric-label">Budget Gap</div><div class="small-metric-value {'negative' if row['budget_gap'] < 0 else 'positive'}">{money_signed(row['budget_gap'])}</div></div>
            <div class="small-metric"><div class="small-metric-label">Gap %</div><div class="small-metric-value {'negative' if row['budget_gap_pct'] < 0 else 'positive'}">{pct(row['budget_gap_pct'])}</div></div>
            <div class="small-metric"><div class="small-metric-label">Contribution to Total Sales</div><div class="small-metric-value">{pct(row['net_sales_mix_pct'])}</div></div>
            <div class="small-metric"><div class="small-metric-label">ATV</div><div class="small-metric-value">{money(row['avg_sales_price'])}</div></div>
        </div>

        <div class="department-subsection-heading {css_class}-subsection-heading">
            {section_icon(dept_name, "subsection-heading-icon")}
            <div>
                <div class="department-subsection-title">Top 5 Brands by Subdepartment</div>
                <div class="department-subsection-note">Ranked by net sales within each subdepartment</div>
            </div>
        </div>
        <div class="brand-grid">{brand_columns}</div>
        {seller_html}
    </div>
    """



def value_to_float(value, default=0.0):
    try:
        if value is None:
            return default
        return float(value)
    except Exception:
        return default


def format_spark_cell(value, display_text=None, kind="gap", max_abs=1.0):
    """Compact inline bar used inside tables.

    The old table stacked the bar above the number, which made each row heavy.
    This version keeps a short spark bar beside the value so the table remains scannable.
    """
    value = value_to_float(value)
    max_abs = max(value_to_float(max_abs, 1.0), 0.000001)
    width = min(abs(value) / max_abs * 100, 100)
    text = display_text if display_text is not None else str(value)

    if kind in {"footwear", "apparel", "accessories"}:
        css = f"spark-{kind}"
    elif kind == "contribution":
        css = "spark-contribution"
    else:
        css = "spark-positive" if value >= 0 else "spark-negative"

    value_class = "positive" if value > 0 else "negative" if value < 0 else "neutral-text"

    return f"""
    <div class="spark-cell">
        <div class="spark-track"><div class="spark-fill {css}" style="width:{width:.1f}%"></div></div>
        <span class="spark-value {value_class}">{text}</span>
    </div>
    """


def format_colored_value(value, display_text=None):
    value = value_to_float(value)
    text = display_text if display_text is not None else str(value)
    value_class = "positive" if value > 0 else "negative" if value < 0 else "neutral-text"
    return f'<span class="value-text {value_class}">{text}</span>'


def build_subdepartment_gap_table(subdept):
    display = subdept.copy()
    columns = [
        "group", "department", "net_sales", "budget", "budget_gap",
        "budget_gap_pct", "mix_pct", "units", "avg_sales_price"
    ]
    display = display[[c for c in columns if c in display.columns]].copy()

    max_gap = display["budget_gap"].abs().max() if "budget_gap" in display.columns and not display.empty else 1
    max_contribution = display["mix_pct"].abs().max() if "mix_pct" in display.columns and not display.empty else 1

    rows = []
    for _, row in display.iterrows():
        css = department_class(row.get("group", ""))
        rows.append(f"""
        <tr class="{css}-gap-row">
            <td class="gap-group-cell"><span class="gap-group-label">{section_icon(row.get('group', ''), 'gap-group-icon')}{row.get('group', '')}</span></td>
            <td class="text-left">{subsection_icon(row.get('department', ''))}<span class="subdepartment-name">{row.get('department', '')}</span></td>
            <td>{money(row.get('net_sales', 0))}</td>
            <td>{money(row.get('budget', 0))}</td>
            <td>{format_spark_cell(row.get('budget_gap', 0), money_signed(row.get('budget_gap', 0)), kind='gap', max_abs=max_gap)}</td>
            <td>{format_colored_value(row.get('budget_gap_pct', 0), pct(row.get('budget_gap_pct', 0)))}</td>
            <td>{format_spark_cell(row.get('mix_pct', 0), pct(row.get('mix_pct', 0)), kind='contribution', max_abs=max_contribution)}</td>
            <td>{number(row.get('units', 0))}</td>
            <td>{money(row.get('avg_sales_price', 0))}</td>
        </tr>
        """)

    return f"""
    <table class="data-table visual-table compact-visual-table">
        <thead>
            <tr>
                <th>Group</th>
                <th>Subdepartment</th>
                <th>Net Sales</th>
                <th>Budget</th>
                <th>Budget Gap</th>
                <th>Gap %</th>
                <th>Contribution to Total Sales</th>
                <th>Units</th>
                <th>ATV</th>
            </tr>
        </thead>
        <tbody>{''.join(rows)}</tbody>
    </table>
    """


def svg_icon(name):
    paths = {
        "sales": '<path d="M4 19V9m6 10V5m6 14v-7m4 7H2"/>',
        "budget": '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="3"/><path d="M12 2v3m0 14v3m10-10h-3M5 12H2"/>',
        "gap": '<path d="M4 18l5-5 4 4 7-9"/><path d="M15 8h5v5"/>',
        "conversion": '<path d="M20 7h-9a4 4 0 0 0-4 4v1"/><path d="M17 4l3 3-3 3"/><path d="M4 17h9a4 4 0 0 0 4-4v-1"/><path d="M7 20l-3-3 3-3"/>',
        "traffic": '<circle cx="8" cy="8" r="3"/><circle cx="16" cy="8" r="3"/><path d="M2 20v-2a5 5 0 0 1 10 0v2m0 0v-2a5 5 0 0 1 10 0v2"/>',
        "transactions": '<path d="M6 3h12v18l-3-2-3 2-3-2-3 2z"/><path d="M9 8h6m-6 4h6"/>',
        "atv": '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 10h18m4-3h4"/>',
        "upt": '<path d="M4 7l8-4 8 4-8 4z"/><path d="M4 7v10l8 4 8-4V7m-8 4v10"/>',
    }
    return f'<svg class="kpi-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">{paths.get(name, paths["sales"])}</svg>'


def _trend_note(note, tone="neutral"):
    """Add a compact direction marker to comparison notes without changing the data."""
    text = str(note or "")
    if tone == "positive" and re.search(r"\d", text):
        return f'<span class="trend-arrow">▲</span> {text}'
    if tone == "negative" and re.search(r"\d", text):
        return f'<span class="trend-arrow">▼</span> {text}'
    return text


def kpi_card(title, value, note, icon, progress=None, progress_label="", tone="neutral"):
    note_class = f"kpi-note-{tone}"

    if progress is None:
        utility = '<div class="kpi-empty-space" aria-hidden="true"></div>'
    else:
        width = max(0, min(float(progress) * 100, 100))
        # Keep the metric label on the left and its number on the right so the
        # progress-bar information can be scanned in one horizontal line.
        match = re.match(r"\s*([+-]?\d+(?:\.\d+)?%)\s*(.*)", str(progress_label or ""))
        if match:
            progress_value = match.group(1)
            progress_text = match.group(2).strip() or "Progress"
        else:
            progress_value = ""
            progress_text = str(progress_label or "Progress")

        utility = (
            f'<div class="kpi-progress-meta">'
            f'<span class="kpi-progress-text">{progress_text}</span>'
            f'<span class="kpi-progress-value">{progress_value}</span>'
            f'</div>'
            f'<div class="kpi-progress">'
            f'<div class="kpi-progress-fill {tone}" style="width:{width:.1f}%"></div>'
            f'</div>'
        )

    return (
        f'<div class="card tone-{tone}">'
        f'<div class="card-main-row">'
        f'<div class="card-value">{value}</div>'
        f'<div class="card-identity">'
        f'<span class="icon-shell">{svg_icon(icon)}</span>'
        f'<div class="card-title">{title}</div>'
        f'</div>'
        f'</div>'
        f'<div class="card-note {note_class}">{_trend_note(note, tone)}</div>'
        f'<div class="card-utility">{utility}</div>'
        f'</div>'
    )


def build_top_sellers_table(df, top_n=None, show_context=False, compact=False):
    if df is None or df.empty:
        return '<div class="empty-note">No product sales were available.</div>'
    work=df.copy().sort_values(["rank","net_sales"] if "rank" in df.columns else ["net_sales"], ascending=[True,False] if "rank" in df.columns else [False])
    if top_n: work=work.head(top_n)
    rows=[]
    for _,r in work.iterrows():
        context = f'<span class="product-context">{r.get("group","")} · {r.get("department","")}</span>' if show_context else ''
        stock=float(r.get("units_on_hand",0) or 0); stock_class='stock-low' if stock <= 3 else ''
        rows.append(f'<tr><td class="rank-cell">{int(r.get("rank",0))}</td><td class="product-cell"><strong>{r.get("product","")}</strong><span>{r.get("brand","")} · {r.get("style_no","")}</span>{context}</td><td>{money(r.get("net_sales",0))}</td><td>{number(r.get("units_sold",0))}</td><td class="{stock_class}">{number(stock)}</td></tr>')
    compact_class = " compact-product-table" if compact else ""
    return f'<table class="data-table product-table{compact_class}"><thead><tr><th>#</th><th>Product</th><th>Sales</th><th>Units</th><th>Stock</th></tr></thead><tbody>'+''.join(rows)+'</tbody></table>'


def build_grouped_top_sellers(df, level="group"):
    if df is None or df.empty:
        return '<div class="empty-note">No top-seller data available.</div>'

    work = df.copy()
    work["_group_order"] = work["group"].apply(department_sort_key)
    if "department" in work.columns:
        work["_sub_order"] = work["department"].apply(subdepartment_sort_key)
    else:
        work["_sub_order"] = 99
    sort_cols = ["_group_order", "_sub_order"]
    if "rank" in work.columns:
        sort_cols.append("rank")
    work = work.sort_values(sort_cols)

    if level == "group":
        blocks = []
        for group_name, g in work.groupby("group", sort=False):
            css = department_class(group_name)
            blocks.append(
                f'<div class="seller-block {css}-seller">'
                f'<div class="seller-block-title">{section_icon(group_name, "seller-title-icon")}<span>{group_name}</span></div>'
                f'{build_top_sellers_table(g)}</div>'
            )
        return '<div class="seller-grid">' + ''.join(blocks) + '</div>'

    sections = []
    for group_name, group_df in work.groupby("group", sort=False):
        css = department_class(group_name)
        blocks = []
        for department_name, g in group_df.groupby("department", sort=False):
            blocks.append(
                f'<div class="seller-block {css}-seller">'
                f'<div class="seller-block-title">{subsection_icon(department_name)}<span>{group_name} · {department_name}</span></div>'
                f'{build_top_sellers_table(g)}</div>'
            )
        sections.append(
            f'<div class="seller-department-section {css}-seller-section">'
            f'<div class="seller-department-heading">{section_icon(group_name, "seller-heading-icon")}<span>{group_name}</span></div>'
            f'<div class="seller-grid">{"".join(blocks)}</div></div>'
        )
    return ''.join(sections)


def build_html_report(
    data_model=DATA_MODEL,
    store_name="Store",
    report_period="Selected period",
    generated_date=None,
    report_title="KPI Performance Report",
):
    from datetime import date

    if generated_date is None:
        generated_date = date.today().isoformat()

    outputs = build_report_calculations(data_model, store_name=store_name)

    kpi = outputs["kpi_master"]
    dept = outputs["department_kpis"].copy()
    brands = outputs["top_brands_overall"].copy()
    top_brands_fixed = outputs["top_brands_fixed"].copy()
    subdept = outputs["sub_department_kpis"].copy()
    opportunities = outputs["opportunity_ranking"].copy()
    validation = outputs["validation"].copy()
    summary = highlight_summary(outputs["executive_summary_text"])
    top_sellers_global = outputs.get("top_sellers_global")
    top_sellers_department = outputs.get("top_sellers_department")
    top_sellers_subdepartment = outputs.get("top_sellers_subdepartment")

    # Enforce one visual order throughout the report.
    if not dept.empty:
        dept["_order"] = dept["group"].apply(department_sort_key)
        dept = dept.sort_values("_order").drop(columns="_order")
    if not subdept.empty:
        subdept["_group_order"] = subdept["group"].apply(department_sort_key)
        subdept["_sub_order"] = subdept["department"].apply(subdepartment_sort_key)
        subdept = subdept.sort_values(["_group_order", "_sub_order"]).drop(columns=["_group_order", "_sub_order"])

    budget_df = dept.rename(columns={"group": "Department", "budget_gap": "Budget Gap"})
    sales_mix_df = dept.rename(columns={
        "group": "Department",
        "net_sales": "Net Sales",
        "net_sales_mix_pct": "Sales %",
    })
    if "Sales %" in sales_mix_df.columns and sales_mix_df["Sales %"].max() <= 1:
        sales_mix_df["Sales %"] = sales_mix_df["Sales %"] * 100

    brand_df = brands.rename(columns={"brand": "Brand", "net_sales": "Sales", "group": "Group"})
    make_budget_gap_chart(budget_df)
    make_sales_mix_donut(sales_mix_df)
    make_top_brands_chart(brand_df)

    make_subdepartment_mix_chart(subdept, "Footwear", "footwear_mix.png")
    make_subdepartment_mix_chart(subdept, "Apparel", "apparel_mix.png")
    make_subdepartment_mix_chart(subdept, "Accessories", "accessories_mix.png")

    net_sales = get_kpi_value(kpi, "Net Sales")
    budget = get_kpi_value(kpi, "Budget")
    budget_gap = get_kpi_value(kpi, "Budget Gap")
    conversion = get_kpi_value(kpi, "Conversion %")
    footfall = get_kpi_value(kpi, "Footfall")
    transactions = get_kpi_value(kpi, "Transactions")
    atv = get_kpi_value(kpi, "ATV")
    ipc = get_kpi_value(kpi, "IPC / UPT")
    sales_vs_ly = get_kpi_value(kpi, "Sales vs LY", "variance_pct")
    budget_gap_pct = get_kpi_value(kpi, "Budget Gap %")
    budget_attainment = get_kpi_value(kpi, "Budget Attainment %")
    transactions_vs_ly = get_kpi_value(kpi, "Transactions", "variance_pct")
    ipc_vs_ly = get_kpi_value(kpi, "IPC / UPT", "variance_pct")

    department_table = build_department_table(dept)
    top_brands_table = build_top_brands_table(brands)
    subdepartment_gap_table = build_subdepartment_gap_table(subdept)
    opportunities_html = build_opportunities_table(opportunities)
    validation_html = build_validation_table(validation)
    global_sellers_html = build_top_sellers_table(top_sellers_global, show_context=True)
    department_sellers_html = build_grouped_top_sellers(top_sellers_department, "group")
    subdepartment_sellers_html = build_grouped_top_sellers(top_sellers_subdepartment, "subdepartment")

    department_cards = "\n".join(
        build_department_card(row, top_brands_fixed, top_sellers_department)
        for _, row in dept.iterrows()
    )

    html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{store_name} KPI Report</title>
    <style>
        @page {{ size: Letter; margin: 10mm; }}
        @media print {{
            body {{ background: #fff !important; }}
            .page {{ box-shadow: none !important; margin: 0 auto !important; }}
            .page-break {{ break-before: page; page-break-before: always; }}
            * {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        }}
        * {{ box-sizing: border-box; }}
        body {{ font-family: Arial, Helvetica, sans-serif; margin: 0; background: #efefef; color: #202020; }}
        .page {{ width: 1200px; margin: 18px auto; background: white; padding: 22px 24px; box-shadow: 0 4px 22px rgba(0,0,0,0.12); }}
        .page-break {{ page-break-before: always; }}
        .header {{ background: #0F172A; color: white; padding: 20px 24px; margin-bottom: 16px; display: flex; justify-content: space-between; align-items: flex-end; }}
        .title {{ font-size: 34px; font-weight: 800; line-height: 1.1; }}
        .subtitle {{ font-size: 16px; margin-top: 8px; opacity: 0.9; }}
        .header-meta {{ text-align: right; font-size: 14px; line-height: 1.5; }}
        .header-actions {{ display:flex; flex-direction:column; align-items:flex-end; gap:9px; }}
        .print-button {{
            appearance:none;
            border:1px solid rgba(255,255,255,.45);
            background:rgba(255,255,255,.10);
            color:#FFFFFF;
            border-radius:7px;
            padding:7px 11px;
            font-size:11px;
            font-weight:800;
            letter-spacing:.01em;
            cursor:pointer;
            display:inline-flex;
            align-items:center;
            gap:6px;
        }}
        .print-button:hover {{ background:rgba(255,255,255,.18); }}
        .print-button:focus {{ outline:2px solid rgba(255,255,255,.65); outline-offset:2px; }}

        .cards {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:8px; margin-bottom:8px; align-items:stretch; }}
        .card {{ border-radius:11px; display:flex; flex-direction:column; border:1px solid #D9E2EC; padding:11px 13px 9px; background:linear-gradient(180deg,#FFFFFF 0%,#FBFCFE 100%); min-height:108px; height:108px; box-shadow:0 3px 10px rgba(15,23,42,.05); position:relative; overflow:hidden; box-sizing:border-box; }}
        .card::before {{ content:""; position:absolute; inset:0 auto 0 0; width:2.5px; background:#CBD5E1; opacity:.9; }}
        .card.tone-positive::before {{ background:#238A4B; }} .card.tone-negative::before {{ background:#C63D3D; }} .card.tone-neutral::before {{ background:#356E9D; }}
        .card-title {{ font-size:11.5px; line-height:1.1; text-transform:uppercase; font-weight:800; color:#50555B; letter-spacing:.035em; text-align:right; max-width:106px; }}
        .card-value {{ font-size:27px; font-weight:800; line-height:1; letter-spacing:-.025em; color:#16181C; white-space:nowrap; }}
        .card-note {{ font-size:10px; line-height:1.15; margin-top:5px; font-weight:750; min-height:12px; background:transparent!important; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
        .negative {{ color:#D00000; }}
        .positive {{ color:#16803A; }}
        .kpi-note-positive {{ color:#16803A; }}
        .kpi-note-negative {{ color:#C62828; }}
        .kpi-note-neutral {{ color:#667085; font-weight:600; }}
        .chart-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 12px; margin-bottom: 12px; }}
        .single-chart-row {{ grid-template-columns: 1fr 1fr; }}
        .chart-card {{ background: white; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px; }}
        .chart-card h3 {{ margin: 0 0 8px 0; font-size: 17px; font-weight: 800; display:flex; align-items:center; gap:7px; }} .chart-heading-icon {{ width:21px; height:21px; }}
        .chart-card img {{ width: 100%; display: block; }}
        .section {{ margin-top: 18px; }}
        .section-title {{ background: #0F172A; color: white; padding: 9px 12px; font-size: 15px; font-weight: 800; margin-bottom: 10px; border-radius: 6px; }}
        .summary {{ font-size: 16px; line-height: 1.65; background: linear-gradient(90deg,#F1F5F9,#FAFCFE); border-left: 5px solid #315F86; padding: 16px 20px; box-shadow: inset 0 0 0 1px #E5EAF0; }}
        .summary-chip {{ display:inline-block; padding:1px 6px; border-radius:999px; font-weight:800; line-height:1.35; margin:0 1px; border:1px solid transparent; }}
        .money-chip {{ background:#E9EEF5; color:#243B53; border-color:#CBD5E1; }}
        .positive-chip {{ background:#E7F4EB; color:#176B35; border-color:#A8D5B5; }}
        .negative-chip {{ background:#FBEAEA; color:#A82424; border-color:#E5B5B5; }}
        .mix-chip {{ background:#E8F0F7; color:#24587D; border-color:#B8CBDA; }}
        .note-box {{ font-size: 13px; color: #444; background: #f7f7f7; border-left: 4px solid #555; padding: 12px 14px; margin: 12px 0; }}
        .table-grid {{ display: grid; grid-template-columns: 1.35fr 1fr; gap: 24px; }}
        .subsection-label {{ font-size: 13px; font-weight: 800; color: #333; margin: 0 0 8px 0; text-transform: uppercase; letter-spacing: 0.03em; }}
        .data-table {{ width: 100%; border-collapse: collapse; font-size: 13px; background: white; }}
        .data-table th {{ background: #1E293B; color: white; padding: 8px 7px; text-align: center; font-weight: 800; }}
        .data-table td {{ border-bottom: 1px solid #E2E8F0; padding: 7px 7px; text-align: center; vertical-align: middle; }}
        .data-table tr:nth-child(even) {{ background: #f8f8f8; }}
        .brand-performance-table th:first-child, .brand-performance-table td:first-child {{ text-align: left; }}
        .brand-performance-table th:nth-child(2), .brand-performance-table td:nth-child(2) {{ text-align: left; font-weight: 700; }}
        .opportunities-table td:last-child {{ text-align: left; }}
        .visual-table th {{ background: #005A64; color: white; }}
        .visual-table td {{ padding: 10px 9px; }}
        .compact-visual-table {{ font-size: 12.5px; }}
        .compact-visual-table .text-left {{ text-align: left; }}
        .spark-cell {{ display: grid; grid-template-columns: 96px 68px; align-items: center; gap: 10px; min-width: 170px; }}
        .spark-track {{ width: 96px; height: 8px; background: #edf0f2; border-radius: 999px; overflow: hidden; }}
        .spark-fill {{ height: 100%; border-radius: 999px; }}
        .spark-positive {{ background: #3FA66B; }}
        .spark-negative {{ background: #E05A5A; }}
        .spark-contribution {{ background: #5B8DB8; }}
        .spark-value {{ font-weight: 800; font-size: 12px; text-align: right; white-space: nowrap; }}
        .value-text {{ font-weight: 800; white-space: nowrap; }}
        .neutral-text {{ color: #333; }}
        .department-grid {{ display: grid; grid-template-columns: 1fr; gap: 16px; }}
        .department-card {{ border: 1px solid #E2E8F0; border-radius: 10px; padding: 13px; background: #fff; }}
        .department-header {{ color: white; padding: 12px 16px; border-radius: 10px; font-size: 23px; font-weight: 800; margin-bottom: 14px; display:flex; align-items:center; gap:10px; box-shadow: inset 0 -2px 0 rgba(0,0,0,.12); }}
        .section-icon {{ width:30px; height:30px; flex:0 0 auto; }}
        .subsection-icon {{ width:20px; height:20px; flex:0 0 auto; }}
        .footwear {{ background: #176FA6; }}
        .apparel {{ background: #D96A2B; }}
        .accessories {{ background: #4A9455; }}
        .neutral {{ background: #555; }}
        .metric-row {{ display: grid; grid-template-columns: repeat(6, 1fr); gap: 10px; margin-bottom: 18px; }}
        .small-metric {{ background: #f7f7f7; border-radius: 10px; padding: 11px; text-align: center; border: 1px solid #e2e2e2; }}
        .small-metric-label {{ font-size: 11px; font-weight: 800; color: #555; text-transform: uppercase; }}
        .small-metric-value {{ font-size: 18px; font-weight: 800; margin-top: 5px; }}
        .brand-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }}
        .brand-column {{ border: 1px solid #ddd; border-radius: 10px; overflow: hidden; background: #fff; }}
        .brand-column-title {{ color: white; padding: 8px 11px; font-size: 13px; font-weight: 800; display:flex; justify-content:center; align-items:center; gap:7px; }}
        .footwear-shade-1 .brand-column-title {{ background:#155F8E; }} .footwear-shade-2 .brand-column-title {{ background:#2879A8; }} .footwear-shade-3 .brand-column-title {{ background:#3D8DB8; }} .footwear-shade-4 .brand-column-title {{ background:#559FC4; }} .footwear-shade-5 .brand-column-title {{ background:#5FA3C5; color:#FFFFFF; }}
        .apparel-shade-1 .brand-column-title {{ background:#B9531F; }} .apparel-shade-2 .brand-column-title {{ background:#D5682B; }} .apparel-shade-3 .brand-column-title {{ background:#E67D43; }} .apparel-shade-4 .brand-column-title {{ background:#EC9563; }} .apparel-shade-5 .brand-column-title {{ background:#E99165; color:#FFFFFF; }}
        .accessories-shade-1 .brand-column-title {{ background:#34753F; }} .accessories-shade-2 .brand-column-title {{ background:#4A9455; }} .accessories-shade-3 .brand-column-title {{ background:#64A96D; }} .accessories-shade-4 .brand-column-title {{ background:#7DB985; }} .accessories-shade-5 .brand-column-title {{ background:#78AE7E; color:#FFFFFF; }}
        .mini-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
        .mini-table th {{ background: #f1f1f1; color: #222; padding: 7px; font-weight: 800; border-bottom: 1px solid #ddd; }}
        .mini-table td {{ padding: 7px; border-bottom: 1px solid #eee; text-align: center; }}
        .mini-table tr:nth-child(even) {{ background: #fafafa; }}
        .empty-note {{ padding: 14px; text-align: center; color: #777; font-size: 12px; }}
        .footer {{ margin-top: 16px; font-size: 11px; color: #666; text-align: right; }}

        .card-main-row {{ display:flex; align-items:flex-start; justify-content:space-between; gap:10px; }}
        .card-identity {{ margin-left:auto; display:flex; align-items:center; justify-content:flex-end; gap:7px; min-width:0; }}
        .icon-shell {{ width:26px; height:26px; display:inline-flex; align-items:center; justify-content:center; color:#28658F; flex:0 0 26px; }}
        .kpi-icon {{ width:22px; height:22px; color:currentColor; flex:0 0 auto; stroke-width:1.9; }}
        .card-utility {{ margin-top:auto; padding-top:5px; min-height:27px; }}
        .kpi-progress-meta {{ display:flex; align-items:baseline; justify-content:space-between; gap:8px; margin-bottom:3px; }}
        .kpi-progress-text {{ color:#64748B; font-size:9.6px; line-height:1.1; font-weight:650; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
        .kpi-progress-value {{ color:#334155; font-size:10.5px; line-height:1.1; font-weight:850; white-space:nowrap; }}
        .kpi-progress {{ height:8px; background:#E7ECF1; border-radius:999px; overflow:hidden; box-shadow:inset 0 0 0 1px rgba(100,116,139,.08); }}
        .kpi-progress-fill {{ height:100%; border-radius:99px; background:#64748B; }}
        .kpi-progress-fill.positive {{ background:#10B981; }} .kpi-progress-fill.negative {{ background:#EF4444; }} .kpi-progress-fill.neutral {{ background:#1D4ED8; }}
        .trend-arrow {{ font-size:8px; vertical-align:1px; margin-right:1px; }}
        .kpi-empty-space {{ height:17px; }}
        .seller-grid {{ display:grid; grid-template-columns:repeat(2,1fr); gap:12px; }}
        .seller-block {{ border:1px solid #E2E8F0; border-radius:9px; overflow:hidden; break-inside:avoid; }}
        .seller-block-title {{ padding:8px 10px; font-weight:800; display:flex; align-items:center; gap:7px; border-bottom:1px solid #D7E0EA; }}
        .seller-title-icon {{ width:18px; height:18px; }}
        .footwear-seller {{ border-top:4px solid #176FA6; }} .footwear-seller .seller-block-title {{ background:#EAF3F8; color:#174A68; }}
        .apparel-seller {{ border-top:4px solid #D96A2B; }} .apparel-seller .seller-block-title {{ background:#FBEEE7; color:#7D391B; }}
        .accessories-seller {{ border-top:4px solid #4A9455; }} .accessories-seller .seller-block-title {{ background:#EDF6EF; color:#285F31; }}
        .product-table {{ font-size:11.5px; }} .product-table th:first-child {{ width:34px; }}
        .product-cell {{ text-align:left!important; min-width:230px; }} .product-cell span {{ display:block; color:#64748B; font-size:10.5px; margin-top:2px; }}
        .product-context {{ color:#1D4ED8!important; }} .rank-cell {{ font-weight:800; color:#1D4ED8; }} .stock-low {{ color:#B91C1C; font-weight:900; text-decoration:underline; text-decoration-thickness:2px; }}
        .spark-footwear {{ background:#176FA6; }} .spark-apparel {{ background:#D96A2B; }} .spark-accessories {{ background:#4A9455; }}
        .department-name-cell {{ min-width:126px; text-align:left!important; }}
        .department-pill {{ display:inline-flex; align-items:center; gap:7px; color:white; border-radius:999px; padding:4px 9px; font-weight:800; font-size:11px; width:112px; justify-content:flex-start; }}
        .table-dept-icon {{ width:16px; height:16px; flex:0 0 16px; }}
        .department-performance-table td {{ padding-top:8px; padding-bottom:8px; }}
        .department-performance-table .spark-cell {{ grid-template-columns:72px 62px; min-width:142px; gap:7px; }}
        .department-performance-table .spark-track {{ width:72px; }}
        .footwear-row {{ background:#F7FBFD!important; }} .apparel-row {{ background:#FFF9F5!important; }} .accessories-row {{ background:#F7FBF8!important; }}

        .embedded-sellers {{ margin-top:14px; border:1px solid #DCE4EC; border-radius:9px; overflow:hidden; }}
        .embedded-sellers-title {{ display:flex; align-items:center; gap:8px; padding:8px 11px; font-size:13px; font-weight:800; }}
        .footwear-department-card {{ border-top:4px solid #176FA6; }}
        .apparel-department-card {{ border-top:4px solid #D96A2B; }}
        .accessories-department-card {{ border-top:4px solid #4A9455; }}
        .footwear-department-card .embedded-sellers-title {{ background:#EAF3F8; color:#174A68; }}
        .apparel-department-card .embedded-sellers-title {{ background:#FBEEE7; color:#7D391B; }}
        .accessories-department-card .embedded-sellers-title {{ background:#EDF6EF; color:#285F31; }}

        .footwear-seller .rank-cell, .footwear-seller .product-context {{ color:#176FA6!important; }}
        .apparel-seller .rank-cell, .apparel-seller .product-context {{ color:#D96A2B!important; }}
        .accessories-seller .rank-cell, .accessories-seller .product-context {{ color:#4A9455!important; }}
        .seller-department-section {{ margin:14px 0 20px; }}
        .seller-department-heading {{ display:flex; align-items:center; gap:9px; padding:9px 12px; border-radius:7px; font-weight:800; margin-bottom:10px; color:white; }}
        .seller-heading-icon {{ width:22px; height:22px; }}
        .footwear-seller-section .seller-department-heading {{ background:#176FA6; }}
        .apparel-seller-section .seller-department-heading {{ background:#D96A2B; }}
        .accessories-seller-section .seller-department-heading {{ background:#4A9455; }}

        .gap-group-cell {{ text-align:left!important; }}
        .gap-group-label {{ display:inline-flex; align-items:center; gap:6px; font-weight:700; }}
        .gap-group-icon {{ width:17px; height:17px; }}
        .compact-visual-table td.text-left {{ display:flex; align-items:center; gap:7px; }}
        .subdepartment-name {{ white-space:nowrap; }}
        .footwear-gap-row {{ background:#F5FAFD!important; }}
        .apparel-gap-row {{ background:#FFF7F2!important; }}
        .accessories-gap-row {{ background:#F5FAF6!important; }}
        .footwear-gap-row:nth-child(even) {{ background:#EDF6FB!important; }}
        .apparel-gap-row:nth-child(even) {{ background:#FDF0E8!important; }}
        .accessories-gap-row:nth-child(even) {{ background:#EDF7EF!important; }}

        .chart-card.compact-chart {{ padding:10px 12px 6px; }}
        .chart-card.compact-chart img {{ max-height:335px; object-fit:contain; }}
        .top-brands-chart img {{ max-height:430px; object-fit:contain; }}


        .store-insights-section {{ margin-top:12px; }}
        .store-insights-grid {{
            display:grid;
            grid-template-columns:minmax(0, 1.22fr) minmax(0, 1fr);
            gap:14px;
            align-items:stretch;
        }}
        .store-insight-panel {{
            min-width:0;
            display:flex;
            flex-direction:column;
        }}
        .store-insight-panel .section-title {{ margin-bottom:8px; }}
        .compact-note {{
            min-height:38px;
            display:flex;
            align-items:center;
            padding:9px 12px;
            margin-bottom:9px;
            font-size:11.5px;
            line-height:1.3;
        }}
        .compact-product-table {{
            table-layout:fixed;
            font-size:10.6px;
        }}
        .compact-product-table th,
        .compact-product-table td {{
            padding-top:6px;
            padding-bottom:6px;
        }}
        .compact-product-table th:first-child {{ width:28px; }}
        .compact-product-table th:nth-child(3) {{ width:74px; }}
        .compact-product-table th:nth-child(4) {{ width:50px; }}
        .compact-product-table th:nth-child(5) {{ width:50px; }}
        .compact-product-table .product-cell {{
            min-width:0;
            line-height:1.16;
        }}
        .compact-product-table .product-cell strong {{
            font-size:10.7px;
            line-height:1.16;
        }}
        .compact-product-table .product-cell span {{
            display:inline;
            font-size:9.4px;
            margin-top:0;
        }}
        .compact-product-table .product-context::before {{
            content:" · ";
            color:#94A3B8;
        }}
        .compact-side-chart {{
            flex:1;
            min-height:0;
            display:flex;
            align-items:center;
            justify-content:center;
            padding:7px 7px 3px;
        }}
        .compact-side-chart img {{
            width:100%;
            max-height:382px;
            object-fit:contain;
        }}
        @media (max-width:950px) {{
            .store-insights-grid {{ grid-template-columns:1fr; }}
            .compact-side-chart img {{ max-height:325px; }}
        }}


        .department-subsection-heading {{
            display:flex;
            align-items:center;
            gap:9px;
            margin:14px 0 9px;
            padding:8px 11px;
            border-radius:8px;
            border:1px solid #DCE4EC;
        }}
        .department-subsection-title {{
            font-size:13px;
            line-height:1.15;
            font-weight:800;
        }}
        .department-subsection-note {{
            margin-top:2px;
            font-size:9.5px;
            line-height:1.2;
            color:#64748B;
            font-weight:500;
        }}
        .subsection-heading-icon {{
            width:22px;
            height:22px;
            flex:0 0 22px;
        }}
        .footwear-subsection-heading {{
            background:#F2F8FC;
            color:#174A68;
            border-left:4px solid #176FA6;
        }}
        .apparel-subsection-heading {{
            background:#FFF6F0;
            color:#7D391B;
            border-left:4px solid #D96A2B;
        }}
        .accessories-subsection-heading {{
            background:#F3F9F4;
            color:#285F31;
            border-left:4px solid #4A9455;
        }}

        @media print {{
            * {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
            .print-button {{ display:none!important; }}
            body {{ background: white; }}
            .page {{ width: auto; margin: 0; box-shadow: none; page-break-after: always; }}
            .card, .chart-card, .department-card, .seller-block, .brand-column {{ box-shadow:none!important; break-inside:avoid; }}
            .store-insights-grid {{ grid-template-columns:1.22fr 1fr!important; gap:10px!important; }}
            .store-insight-panel {{ break-inside:avoid; }}
            .footwear, .footwear .brand-column-title, .footwear-seller {{ filter:grayscale(1); }}
            .apparel, .apparel .brand-column-title, .apparel-seller {{ filter:grayscale(1); }}
            .accessories, .accessories .brand-column-title, .accessories-seller {{ filter:grayscale(1); }}
            .footwear {{ background:#3D3D3D!important; }}
            .apparel {{ background:#707070!important; }}
            .accessories {{ background:#A0A0A0!important; color:#111!important; }}
            .spark-footwear {{ background:#3D3D3D!important; }}
            .spark-apparel {{ background:repeating-linear-gradient(45deg,#656565 0,#656565 4px,#8A8A8A 4px,#8A8A8A 8px)!important; }}
            .spark-accessories {{ background:repeating-linear-gradient(90deg,#8C8C8C 0,#8C8C8C 3px,#B0B0B0 3px,#B0B0B0 7px)!important; }}
            .spark-positive {{ background:#4F4F4F!important; }}
            .spark-negative {{ background:repeating-linear-gradient(45deg,#333 0,#333 3px,#777 3px,#777 7px)!important; }}
            .stock-low::after {{ content:" LOW"; font-size:8px; letter-spacing:.04em; }}
            .summary-chip {{ background:#EEE!important; color:#111!important; border-color:#777!important; }}
        }}
    </style>
</head>
<body>
    <div class="page">
        <div class="header">
            <div><div class="title">{store_name} KPI Report</div><div class="subtitle">{report_title}</div></div>
            <div class="header-actions">
                <div class="header-meta"><div><strong>Period:</strong> {report_period}</div><div><strong>Generated:</strong> {generated_date}</div></div>
                <button class="print-button" type="button" onclick="window.print()" title="Open the browser print dialog and choose Save as PDF">
                    <span aria-hidden="true">⎙</span>
                    <span>Print / Save PDF</span>
                </button>
            </div>
        </div>

        <div class="cards">
            {kpi_card("Net Sales", money(net_sales), f"{pct(sales_vs_ly)} vs LY", "sales", budget_attainment, f"{pct(budget_attainment)} Budget attainment", "positive" if budget_attainment >= 1 else "negative")}
            {kpi_card("Budget", money(budget), "Weekly target", "budget", None, "", "neutral")}
            {kpi_card("Budget Gap", money_signed(budget_gap), f"{pct(budget_gap_pct)} vs budget", "gap", budget_attainment, f"{pct(budget_attainment)} Budget attainment", "positive" if budget_gap >= 0 else "negative")}
            {kpi_card("Conversion", pct(conversion), "Transactions ÷ footfall", "conversion", None, "", "neutral")}
        </div>

        <div class="cards">
            {kpi_card("Footfall", number(footfall), "Store traffic", "traffic", None, "", "neutral")}
            {kpi_card("Transactions", number(transactions), f"{pct(transactions_vs_ly)} vs LY", "transactions", min(max(1 + transactions_vs_ly, 0), 1), f"{pct(transactions_vs_ly)} vs LY", "positive" if transactions_vs_ly >= 0 else "negative")}
            {kpi_card("ATV", money(atv), "Average transaction value", "atv", None, "", "neutral")}
            {kpi_card("IPC / UPT", f"{float(ipc):.2f}", f"{pct(ipc_vs_ly)} vs LY", "upt", min(max(1 + ipc_vs_ly, 0), 1), f"{pct(ipc_vs_ly)} vs LY", "positive" if ipc_vs_ly >= 0 else "negative")}
        </div>

        <div class="chart-row">
            <div class="chart-card compact-chart"><h3>Global Sales Contribution</h3><img src="../assets/charts/sales_mix.png" alt="Global Sales Contribution"></div>
            <div class="chart-card compact-chart"><h3>Budget GAP by Department</h3><img src="../assets/charts/budget_gap.png" alt="Budget GAP by Department"></div>
        </div>

        <div class="section"><div class="section-title">Executive Summary</div><div class="summary">{summary}</div></div>

        <div class="section">
            <div class="section-title">Department Performance</div>
            <div class="table-grid">
                <div>
                    <div class="subsection-label">Department KPI Summary</div>
                    {department_table}
                </div>
                <div>
                    <div class="subsection-label">Brand Performance by Department</div>
                    {top_brands_table}
                </div>
            </div>
        </div>

        <div class="section store-insights-section">
            <div class="store-insights-grid">
                <div class="store-insight-panel sellers-panel">
                    <div class="section-title">Top Sellers — Store</div>
                    <div class="note-box compact-note">Ranked by net sales. Stock highlights turn red at three units or fewer.</div>
                    {global_sellers_html}
                </div>
                <div class="store-insight-panel brands-panel">
                    <div class="section-title">Top Brand Categories by Sales</div>
                    <div class="note-box compact-note">Brand — Department labels keep categories distinct.</div>
                    <div class="chart-card top-brands-chart compact-side-chart"><img src="../assets/charts/top_brands.png" alt="Top Brand Categories by Sales"></div>
                </div>
            </div>
        </div>
        <div class="footer">Generated automatically from Tableau exports and report data model.</div>
    </div>

    <div class="page page-break">
        <div class="header">
            <div><div class="title">Department Detail</div><div class="subtitle">Footwear, Apparel and Accessories Performance</div></div>
            <div class="header-meta"><div><strong>Store:</strong> {store_name}</div><div><strong>Period:</strong> {report_period}</div></div>
        </div>
        <div class="department-grid">{department_cards}</div>


        <div class="section">
            <div class="section-title">Subdepartment Contribution</div>

            <div class="chart-row">
                <div class="chart-card compact-chart"><h3>{section_icon("Footwear", "chart-heading-icon")}<span>Footwear Contribution</span></h3><img src="../assets/charts/footwear_mix.png" alt="Footwear Contribution"></div>
                <div class="chart-card compact-chart"><h3>{section_icon("Apparel", "chart-heading-icon")}<span>Apparel Contribution</span></h3><img src="../assets/charts/apparel_mix.png" alt="Apparel Contribution"></div>
            </div>

            <div class="chart-row single-chart-row">
                <div class="chart-card compact-chart"><h3>{section_icon("Accessories", "chart-heading-icon")}<span>Accessories Contribution</span></h3><img src="../assets/charts/accessories_mix.png" alt="Accessories Contribution"></div>
            </div>
        </div>
        <div class="footer">Department sections generated automatically from Tableau brand exports.</div>
    </div>

    <div class="page page-break">
        <div class="header">
            <div><div class="title">Budget Gap Detail</div><div class="subtitle">Subdepartment performance with visual variance bars</div></div>
            <div class="header-meta"><div><strong>Store:</strong> {store_name}</div><div><strong>Period:</strong> {report_period}</div></div>
        </div>

        <div class="section">
            <div class="section-title">Budget Gap by Subdepartment</div>
            {subdepartment_gap_table}
        </div>

        <div class="section"><div class="section-title">Top Sellers by Subdepartment</div><div class="note-box">Top three products are shown for each Men’s, Women’s, Junior, Children’s/Infants and Accessories subdepartment where data is available.</div>{subdepartment_sellers_html}</div>

        <div class="section">
            <div class="section-title">Top Opportunities Ranking</div>
            <div class="note-box">
                Priority Score weighs budget gap severity by contribution to total sales. Larger and more underperforming subdepartments rank higher.
                Recommendations are rule-based prompts by department type and should be reviewed by the manager before execution.
            </div>
            {opportunities_html}
        </div>

        <div class="section"><div class="section-title">Data Quality / Validation Checks</div>{validation_html}</div>
        <div class="footer">Opportunity ranking generated automatically from budget gap, contribution to total sales and priority score.</div>
    </div>
</body>
</html>
"""

    output_path = OUTPUT_DIR / "final_report.html"
    output_path.write_text(html, encoding="utf-8")
    print(f"Created: {output_path}")
    return output_path


if __name__ == "__main__":
    build_html_report()
