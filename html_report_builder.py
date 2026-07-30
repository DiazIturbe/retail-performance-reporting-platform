from pathlib import Path
import base64
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


def image_data_uri(image_path):
    """Return a self-contained image URI for HTML and PDF output."""
    path = Path(image_path)
    if not path.exists():
        return ""
    suffix = path.suffix.lower().lstrip(".")
    subtype = "jpeg" if suffix in {"jpg", "jpeg"} else suffix
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/{subtype};base64,{encoded}"


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


def cover_section_icon(name):
    """Return a PDF-safe cover icon with an explicit department stroke."""
    color = {
        "footwear": "#176FA6",
        "apparel": "#D96A2B",
        "accessories": "#4A9455",
    }.get(department_class(name), "#A9C3DC")
    return section_icon(name, "cover-dept-icon").replace(
        'stroke="currentColor"', f'stroke="{color}"'
    )


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
    <div class="department-card {css_class}-department-card" data-department="{css_class}">
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
            <td class="subdepartment-cell"><span class="subdepartment-label">{subsection_icon(row.get('department', ''))}<span class="subdepartment-name">{row.get('department', '')}</span></span></td>
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
                f'<div class="seller-block {css}-seller" data-department="{css}">'
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
                f'<div class="seller-block {css}-seller" data-department="{css}">'
                f'<div class="seller-block-title">{subsection_icon(department_name)}<span>{group_name} · {department_name}</span></div>'
                f'{build_top_sellers_table(g)}</div>'
            )
        sections.append(
            f'<div class="seller-department-section {css}-seller-section" data-department="{css}">'
            f'<div class="seller-department-heading">{section_icon(group_name, "seller-heading-icon")}<span>{group_name}</span></div>'
            f'<div class="seller-grid">{"".join(blocks)}</div></div>'
        )
    return ''.join(sections)


def build_html_report(
    data_model=DATA_MODEL,
    store_name="Store",
    report_period="Selected period",
    generated_date=None,
    report_title="Store Performance Report",
):
    from datetime import date

    if generated_date is None:
        generated_date = date.today().isoformat()

    asset_dir = Path(__file__).resolve().parent / "assets"
    logo_path = asset_dir / "ddi_logo_transparent.png"
    if not logo_path.exists():
        logo_path = asset_dir / "ddi_favicon.png"
    logo_uri = image_data_uri(logo_path)
    cover_logo = (
        f'<span class="cover-logo-shell"><img class="cover-logo" src="{logo_uri}" alt="DDI Data Solutions"></span>'
        if logo_uri else ""
    )

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
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{store_name} KPI Report</title>

    <style>
        :root {{
            --ink: #172033;
            --muted: #64748b;
            --line: #dbe3ed;
            --soft: #f5f7fa;
            --panel: #ffffff;
            --navy: #0f2747;
            --blue: #2563eb;
            --blue-soft: #eaf2ff;
            --green: #16845b;
            --green-soft: #eaf8f2;
            --red: #c63f46;
            --red-soft: #fff0f1;
            --amber: #a96308;
            --amber-soft: #fff6df;
            --footwear: #176FA6;
            --footwear-soft: #eef5fb;
            --apparel: #D96A2B;
            --apparel-soft: #fff3ea;
            --accessories: #4A9455;
            --accessories-soft: #edf8f2;
            /* Brighter cover-only accents. The report body keeps the more
               restrained department palette above. */
            --cover-footwear: #22A7E3;
            --cover-apparel: #F2762E;
            --cover-accessories: #68C85F;
            --radius: 12px;
        }}

        * {{
            box-sizing: border-box;
        }}

        html, body {{
            margin: 0;
            padding: 0;
            background: #e8edf3;
            color: var(--ink);
            font-family: "Inter", "Segoe UI", Arial, Helvetica, sans-serif;
            font-size: 11px;
            line-height: 1.38;
        }}

        body {{
            padding: 22px 0;
        }}

        .page {{
            width: 279.4mm;
            min-height: 215.9mm;
            margin: 0 auto 20px;
            padding: 10mm;
            background: #fff;
            border-radius: 4px;
            box-shadow: 0 10px 28px rgba(15, 39, 71, 0.12);
            overflow: hidden;
            position: relative;
        }}

        .cover-page {{
            page: cover;
            padding: 0;
            display: flex;
            min-height: 215.9mm;
            background:
                linear-gradient(120deg, rgba(15,39,71,.98), rgba(28,70,116,.96)),
                #0f2747;
            color: #fff;
        }}

        .cover-accent {{
            position: absolute;
            inset: 0 auto 0 0;
            width: 8mm;
            background: linear-gradient(
                90deg,
                rgba(63, 145, 218, .72) 0%,
                rgba(75, 169, 222, .28) 48%,
                rgba(75, 169, 222, 0) 100%
            );
            pointer-events: none;
        }}

        .cover-content {{
            flex: 1;
            min-width: 0;
            padding: 10mm 13mm 9mm 14mm;
            display: flex;
            flex-direction: column;
        }}

        .cover-topbar {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8mm;
            margin-bottom: 7mm;
        }}

        .cover-eyebrow {{
            text-transform: uppercase;
            letter-spacing: .18em;
            font-size: 9px;
            font-weight: 700;
            opacity: .78;
        }}

        .cover-logo-shell {{
            width: 12mm;
            height: 12mm;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            flex: 0 0 auto;
            padding: 0;
            border-radius: 2.4mm;
            background: transparent;
            box-shadow: none;
        }}

        .cover-logo {{
            display: block;
            width: 100%;
            height: 100%;
            object-fit: contain;
            border-radius: 0;
        }}

        .cover-title {{
            font-size: 29px;
            line-height: 1.02;
            font-weight: 800;
            letter-spacing: -.04em;
            max-width: 128mm;
            margin: 0;
        }}

        .cover-kicker {{
            margin-bottom: 3mm;
            color: #83bff0;
            font-size: 9px;
            font-weight: 700;
            letter-spacing: .16em;
            text-transform: uppercase;
        }}

        .cover-subtitle {{
            margin-top: 4mm;
            color: #b9cce0;
            font-size: 13px;
            font-weight: 500;
        }}

        .cover-store {{
            font-size: 17px;
            font-weight: 600;
            color: #d8e9f8;
            margin-top: 6mm;
        }}

        .cover-rule {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: .7mm;
            width: 54mm;
            height: 3px;
            margin: 7mm 0 9mm;
        }}

        .cover-rule span {{ border-radius: 99px; }}
        .cover-rule .rule-footwear {{ background: var(--cover-footwear); }}
        .cover-rule .rule-apparel {{ background: var(--cover-apparel); }}
        .cover-rule .rule-accessories {{ background: var(--cover-accessories); }}

        .cover-meta-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8mm 14mm;
            max-width: 122mm;
        }}

        .cover-meta-item span {{
            display: block;
            text-transform: uppercase;
            letter-spacing: .12em;
            font-size: 8px;
            color: #a9c3dc;
            margin-bottom: 2mm;
        }}

        .cover-meta-item strong {{
            font-size: 12px;
            color: #fff;
            font-weight: 600;
        }}

        .cover-bottom {{
            margin-top: auto;
            padding-top: 4mm;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            gap: 12mm;
        }}

        .cover-description {{
            max-width: 190mm;
            color: #d5e3ef;
            font-size: 10px;
            line-height: 1.55;
        }}

        .cover-version {{
            white-space: nowrap;
            font-size: 9px;
            color: #a9c3dc;
        }}

        .cover-contents {{
            margin-top: 4mm;
            padding-top: 3mm;
            border-top: 1px solid rgba(213, 227, 239, .22);
            color: #a9c3dc;
            font-size: 8px;
            letter-spacing: .015em;
        }}

        .cover-credit {{
            margin-top: 1.5mm;
            color: #8faac2;
            font-size: 7.5px;
        }}

        .cover-main {{
            display: grid;
            grid-template-columns: minmax(0, 1.12fr) minmax(94mm, .88fr);
            gap: 12mm;
            align-items: stretch;
            min-height: 111mm;
        }}

        .cover-left {{ min-width: 0; }}

        .cover-panels {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 3mm;
            align-items: end;
            padding-top: 6mm;
        }}

        .cover-dept-panel {{
            position: relative;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 5mm 3mm 4mm;
            border: 1px solid rgba(255,255,255,.18);
            border-radius: 5mm;
            background: transparent;
            box-shadow: none;
        }}

        .cover-dept-panel.footwear-panel {{ height: 72mm; border-color: var(--cover-footwear); color: var(--cover-footwear); }}
        .cover-dept-panel.apparel-panel {{ height: 90mm; border-color: var(--cover-apparel); color: var(--cover-apparel); }}
        .cover-dept-panel.accessories-panel {{ height: 108mm; border-color: var(--cover-accessories); color: var(--cover-accessories); }}

        .cover-dept-icon-shell {{
            width: 14mm;
            height: 14mm;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border: 0;
            border-radius: 0;
            background: transparent;
        }}

        .cover-dept-icon {{ width: 9mm; height: 9mm; }}

        .cover-dept-name {{
            margin-top: 4mm;
            font-size: 9px;
            font-weight: 800;
            letter-spacing: .08em;
            text-transform: uppercase;
        }}

        .cover-mini-rule {{
            width: 64%;
            height: 2px;
            margin-top: 4mm;
            border-radius: 99px;
            background: currentColor;
        }}

        .cover-momentum {{
            width: 72%;
            height: 12mm;
            margin-top: auto;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}

        .cover-momentum span {{
            width: 2mm;
            height: 2mm;
            border-radius: 50%;
            background: currentColor;
        }}

        .cover-momentum span:nth-child(1) {{ transform: translateY(3mm); }}
        .cover-momentum span:nth-child(2) {{ transform: translateY(1mm); }}
        .cover-momentum span:nth-child(3) {{ transform: translateY(2mm); }}
        .cover-momentum span:nth-child(4) {{ transform: translateY(-1mm); }}
        .cover-momentum span:nth-child(5) {{ transform: translateY(-3mm); }}

        .cover-mini-bars {{
            width: 100%;
            height: 20mm;
            display: flex;
            align-items: flex-end;
            justify-content: center;
            gap: 1.4mm;
            margin-top: 2mm;
            opacity: .32;
        }}

        .cover-mini-bars span {{ width: 2.6mm; border-radius: 1mm 1mm 0 0; background: currentColor; }}

        .header {{
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 10mm;
            padding-bottom: 4mm;
            margin-bottom: 5mm;
            border-bottom: 1px solid var(--line);
        }}

        .title {{
            font-size: 21px;
            line-height: 1.05;
            font-weight: 800;
            letter-spacing: -.025em;
            color: var(--navy);
        }}

        .subtitle {{
            margin-top: 1.6mm;
            color: var(--muted);
            font-size: 10px;
        }}

        .header-meta {{
            min-width: 58mm;
            text-align: right;
            color: var(--muted);
            font-size: 9px;
            line-height: 1.55;
        }}

        .header-meta strong {{
            color: var(--ink);
            font-weight: 700;
        }}

        .cards {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 3mm;
            margin-bottom: 3mm;
        }}

        .card {{
            min-height: 29mm;
            border: 1px solid var(--line);
            border-radius: var(--radius);
            padding: 3.5mm;
            background: var(--panel);
            position: relative;
            overflow: hidden;
        }}

        .card::before {{
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 3px;
            background: #94a3b8;
        }}

        .tone-positive::before {{ background: var(--green); }}
        .tone-negative::before {{ background: var(--red); }}
        .tone-neutral::before {{ background: #7890aa; }}

        .card-main-row {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 3mm;
        }}

        .card-value {{
            font-size: 18px;
            line-height: 1;
            font-weight: 800;
            letter-spacing: -.025em;
            color: var(--navy);
            white-space: nowrap;
        }}

        .card-identity {{
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 2mm;
            color: var(--muted);
            text-align: right;
        }}

        .icon-shell {{
            width: 7mm;
            height: 7mm;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            background: var(--soft);
            color: var(--navy);
            flex: 0 0 auto;
        }}

        .kpi-icon {{
            width: 4mm;
            height: 4mm;
        }}

        .card-title {{
            font-size: 8px;
            line-height: 1.2;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: .055em;
        }}

        .card-note {{
            margin-top: 3mm;
            font-size: 9px;
            color: var(--muted);
            min-height: 4mm;
        }}

        .kpi-note-positive {{ color: var(--green); }}
        .kpi-note-negative {{ color: var(--red); }}
        .trend-arrow {{ font-size: 8px; }}

        .card-utility {{
            margin-top: 2mm;
        }}

        .kpi-empty-space {{ height: 7px; }}

        .kpi-progress-meta {{
            display: flex;
            justify-content: space-between;
            gap: 2mm;
            color: var(--muted);
            font-size: 7.5px;
            margin-bottom: 1mm;
        }}

        .kpi-progress-value {{ font-weight: 700; color: var(--ink); }}

        .kpi-progress {{
            height: 3px;
            background: #e6ebf1;
            border-radius: 99px;
            overflow: hidden;
        }}

        .kpi-progress-fill {{
            height: 100%;
            border-radius: 99px;
            background: #7890aa;
        }}

        .kpi-progress-fill.positive {{ background: var(--green); }}
        .kpi-progress-fill.negative {{ background: var(--red); }}

        .chart-row {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 4mm;
            margin-top: 4mm;
        }}

        .single-chart-row {{
            grid-template-columns: minmax(0, 1fr);
        }}

        .single-chart-row .chart-card {{
            max-width: 50%;
        }}

        .chart-card {{
            border: 1px solid var(--line);
            background: #fff;
            border-radius: var(--radius);
            padding: 3.5mm;
            min-width: 0;
        }}

        .chart-card h3 {{
            margin: 0 0 2mm;
            font-size: 10px;
            color: var(--navy);
            display: flex;
            align-items: center;
            gap: 2mm;
        }}

        .chart-card img {{
            display: block;
            width: 100%;
            max-height: 58mm;
            object-fit: contain;
        }}

        .compact-chart img {{ max-height: 50mm; }}
        .top-brands-chart img {{ max-height: 58mm; }}

        .section {{
            margin-top: 4mm;
        }}

        .section-title {{
            display: flex;
            align-items: center;
            min-height: 7mm;
            margin-bottom: 2.5mm;
            font-size: 12px;
            font-weight: 800;
            color: var(--navy);
            letter-spacing: -.01em;
            white-space: nowrap;
        }}

        .section-title::before {{
            content: "";
            width: 3px;
            height: 5mm;
            margin-right: 2.2mm;
            border-radius: 2px;
            background: #4f8bc9;
        }}

        .subsection-label {{
            color: var(--muted);
            text-transform: uppercase;
            letter-spacing: .06em;
            font-size: 8px;
            font-weight: 700;
            margin: 0 0 1.5mm;
        }}

        .summary {{
            border: 1px solid #d7e3ef;
            border-left: 4px solid #4f8bc9;
            background: #f7fbff;
            border-radius: 8px;
            padding: 3mm 4mm;
            color: #334155;
            line-height: 1.55;
        }}

        .summary-chip {{
            display: inline-block;
            border-radius: 4px;
            padding: 0 .8mm;
            font-weight: 700;
        }}

        .money-chip {{ background: #e8eff8; color: var(--navy); }}
        .positive-chip {{ background: var(--green-soft); color: var(--green); }}
        .negative-chip {{ background: var(--red-soft); color: var(--red); }}
        .mix-chip {{ background: var(--amber-soft); color: var(--amber); }}

        .table-grid {{
            display: grid;
            grid-template-columns: 1.42fr .9fr;
            gap: 4mm;
            align-items: start;
        }}

        .data-table, .mini-table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            border: 1px solid var(--line);
            border-radius: 8px;
            overflow: hidden;
            background: #fff;
        }}

        .data-table th, .mini-table th {{
            padding: 2mm 1.5mm;
            background: #eef2f6;
            border-bottom: 1px solid var(--line);
            color: #475569;
            text-align: right;
            font-size: 8px;
            line-height: 1.25;
            text-transform: uppercase;
            letter-spacing: .035em;
            white-space: nowrap;
        }}

        .data-table td, .mini-table td {{
            padding: 1.8mm 1.5mm;
            border-bottom: 1px solid #edf1f5;
            color: #334155;
            text-align: right;
            vertical-align: middle;
            font-size: 9px;
        }}

        .data-table tr:last-child td, .mini-table tr:last-child td {{
            border-bottom: 0;
        }}

        .data-table tbody tr:nth-child(even) td {{
            background: #fafbfd;
        }}

        .data-table th:first-child,
        .data-table td:first-child,
        .mini-table th:first-child,
        .mini-table td:first-child,
        .text-left {{
            text-align: left;
        }}

        .brand-performance-table td:first-child {{
            font-weight: 600;
        }}

        .department-pill, .gap-group-label {{
            display: inline-flex;
            align-items: center;
            gap: 1.5mm;
            font-weight: 700;
            white-space: nowrap;
        }}

        .section-icon, .table-dept-icon, .gap-group-icon,
        .subsection-icon, .chart-heading-icon,
        .seller-title-icon, .seller-heading-icon,
        .subsection-heading-icon {{
            width: 4mm;
            height: 4mm;
            flex: 0 0 auto;
        }}

        /* WeasyPrint does not always resolve SVG currentColor consistently. */
        .department-header .section-icon {{ color: #fff !important; stroke: #fff !important; }}
        .department-subsection-heading .subsection-heading-icon {{ color: var(--navy); stroke: var(--navy); }}
        .footwear-seller .subsection-icon,
        .footwear-seller-section .seller-heading-icon {{ color: var(--footwear); stroke: var(--footwear); }}
        .apparel-seller .subsection-icon,
        .apparel-seller-section .seller-heading-icon {{ color: var(--apparel); stroke: var(--apparel); }}
        .accessories-seller .subsection-icon,
        .accessories-seller-section .seller-heading-icon {{ color: var(--accessories); stroke: var(--accessories); }}
        .footwear-gap-row .gap-group-icon {{ color: var(--footwear); stroke: var(--footwear); }}
        .apparel-gap-row .gap-group-icon {{ color: var(--apparel); stroke: var(--apparel); }}
        .accessories-gap-row .gap-group-icon {{ color: var(--accessories); stroke: var(--accessories); }}

        .footwear {{ color: var(--footwear); }}
        .apparel {{ color: var(--apparel); }}
        .accessories {{ color: var(--accessories); }}

        .footwear-row td:first-child {{ border-left: 3px solid var(--footwear); }}
        .apparel-row td:first-child {{ border-left: 3px solid var(--apparel); }}
        .accessories-row td:first-child {{ border-left: 3px solid var(--accessories); }}

        .compact-visual-table th:first-child,
        .compact-visual-table td:first-child {{ width: 17%; }}

        .compact-visual-table th:nth-child(2),
        .compact-visual-table td:nth-child(2) {{
            width: 18%;
            text-align: left;
        }}

        .subdepartment-cell {{
            text-align: left !important;
            padding-left: 2mm !important;
        }}

        .subdepartment-label {{
            display: inline-flex;
            align-items: center;
            justify-content: flex-start;
            gap: 1.4mm;
            min-width: 0;
            font-weight: 500;
        }}

        .subdepartment-name {{ line-height: 1.2; }}

        .footwear-gap-row .subdepartment-cell {{ background: var(--footwear-soft) !important; }}
        .apparel-gap-row .subdepartment-cell {{ background: var(--apparel-soft) !important; }}
        .accessories-gap-row .subdepartment-cell {{ background: var(--accessories-soft) !important; }}

        .positive {{ color: var(--green) !important; font-weight: 700; }}
        .negative {{ color: var(--red) !important; font-weight: 700; }}
        .neutral-text {{ color: var(--muted) !important; }}

        .spark-cell {{
            display: grid;
            grid-template-columns: minmax(16mm, 1fr) auto;
            align-items: center;
            gap: 1.5mm;
            min-width: 29mm;
        }}

        .spark-track {{
            height: 3px;
            border-radius: 99px;
            background: #e8edf2;
            overflow: hidden;
        }}

        .spark-fill {{ height: 100%; border-radius: 99px; }}
        .spark-positive {{ background: var(--green); }}
        .spark-negative {{ background: var(--red); }}
        .spark-footwear {{ background: var(--footwear); }}
        .spark-apparel {{ background: var(--apparel); }}
        .spark-accessories {{ background: var(--accessories); }}
        .spark-contribution {{ background: #738aa2; }}

        .spark-value {{
            white-space: nowrap;
            font-size: 8px;
        }}

        .store-insights-grid {{
            display: grid;
            grid-template-columns: 1.25fr 1fr;
            gap: 4mm;
            align-items: start;
        }}

        .store-insight-panel {{
            min-width: 0;
        }}

        .note-box {{
            border-radius: 7px;
            border: 1px solid #dce5ee;
            background: var(--soft);
            padding: 2.4mm 3mm;
            color: var(--muted);
            font-size: 8.5px;
            line-height: 1.45;
            margin-bottom: 2.5mm;
        }}

        .compact-note {{
            padding: 1.8mm 2.5mm;
        }}

        .seller-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 3mm;
        }}

        .seller-block, .department-card, .brand-column {{
            border: 1px solid var(--line);
            border-radius: 9px;
            background: #fff;
            overflow: hidden;
        }}

        .seller-block-title, .seller-department-heading,
        .embedded-sellers-title, .brand-column-title,
        .department-subsection-heading {{
            display: flex;
            align-items: center;
            gap: 2mm;
            font-weight: 800;
            color: var(--navy);
        }}

        .seller-block-title {{
            padding: 2.4mm 3mm;
            border-bottom: 1px solid var(--line);
            background: #f6f8fb;
        }}

        .seller-block.footwear-seller {{ border-top: 3px solid var(--footwear); }}
        .seller-block.apparel-seller {{ border-top: 3px solid var(--apparel); }}
        .seller-block.accessories-seller {{ border-top: 3px solid var(--accessories); }}

        .footwear-seller .seller-block-title {{ background: var(--footwear-soft); color: var(--footwear); }}
        .apparel-seller .seller-block-title {{ background: var(--apparel-soft); color: var(--apparel); }}
        .accessories-seller .seller-block-title {{ background: var(--accessories-soft); color: var(--accessories); }}

        .footwear-seller-section .seller-department-heading {{ color: var(--footwear); }}
        .apparel-seller-section .seller-department-heading {{ color: var(--apparel); }}
        .accessories-seller-section .seller-department-heading {{ color: var(--accessories); }}

        .seller-block .data-table {{
            border: 0;
            border-radius: 0;
        }}

        .seller-department-section {{
            margin-bottom: 4mm;
        }}

        .seller-department-heading {{
            margin-bottom: 2mm;
            font-size: 12px;
        }}

        .product-table td {{
            padding-top: 1.5mm;
            padding-bottom: 1.5mm;
        }}

        .rank-cell {{
            width: 8mm;
            color: var(--muted);
            font-weight: 700;
        }}

        .product-cell {{
            text-align: left !important;
            min-width: 38mm;
        }}

        .product-cell strong,
        .product-cell span {{
            display: block;
        }}

        .product-cell strong {{
            font-size: 8.8px;
            color: var(--ink);
        }}

        .product-cell span {{
            margin-top: .5mm;
            color: var(--muted);
            font-size: 7.4px;
        }}

        .stock-low {{
            color: var(--red) !important;
            font-weight: 800;
            background: var(--red-soft) !important;
        }}

        .department-grid {{
            display: grid;
            grid-template-columns: 1fr;
            gap: 4mm;
        }}

        .department-card {{
            padding: 0;
        }}

        .department-header {{
            display: flex;
            align-items: center;
            gap: 2mm;
            padding: 2.5mm 3.5mm;
            color: #fff;
            font-weight: 800;
            font-size: 12px;
        }}

        .department-header.footwear {{ background: var(--footwear); }}
        .department-header.apparel {{ background: var(--apparel); }}
        .department-header.accessories {{ background: var(--accessories); }}
        .department-header.neutral {{ background: #64748b; }}

        .metric-row {{
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 2mm;
            padding: 3mm;
            border-bottom: 1px solid var(--line);
        }}

        .small-metric {{
            padding: 2mm;
            border-radius: 7px;
            background: var(--soft);
            min-width: 0;
            min-height: 16mm;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}

        .small-metric-label {{
            color: var(--muted);
            text-transform: uppercase;
            letter-spacing: .035em;
            font-size: 7.2px;
            line-height: 1.15;
            min-height: 0;
        }}

        .small-metric-value {{
            margin-top: 1.5mm;
            font-size: 11px;
            line-height: 1.1;
            font-weight: 800;
            color: var(--navy);
            white-space: nowrap;
        }}

        .department-subsection-heading {{
            padding: 3mm 3mm 1.5mm;
        }}

        .department-subsection-title {{
            font-size: 10.5px;
        }}

        .department-subsection-note {{
            color: var(--muted);
            font-size: 8px;
            font-weight: 400;
        }}

        .brand-grid {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 2.5mm;
            padding: 2mm 3mm 3mm;
        }}

        .brand-column-title {{
            padding: 2mm 2.5mm;
            background: #f6f8fb;
            border-bottom: 1px solid var(--line);
            font-size: 9px;
        }}

        .brand-column .mini-table {{
            border: 0;
            border-radius: 0;
        }}

        .brand-column .mini-table th,
        .brand-column .mini-table td {{
            padding: 1.35mm 1.5mm;
            font-size: 8px;
        }}

        .footwear-column {{ border-top: 3px solid var(--footwear); }}
        .apparel-column {{ border-top: 3px solid var(--apparel); }}
        .accessories-column {{ border-top: 3px solid var(--accessories); }}

        .embedded-sellers {{
            margin: 0 3mm 3mm;
            border: 1px solid var(--line);
            border-radius: 8px;
            overflow: hidden;
        }}

        .embedded-sellers-title {{
            padding: 2mm 2.5mm;
            background: #f6f8fb;
            border-bottom: 1px solid var(--line);
        }}

        .embedded-sellers.footwear-seller {{ border-top: 3px solid var(--footwear); }}
        .embedded-sellers.apparel-seller {{ border-top: 3px solid var(--apparel); }}
        .embedded-sellers.accessories-seller {{ border-top: 3px solid var(--accessories); }}

        .embedded-sellers.footwear-seller .embedded-sellers-title {{ background: var(--footwear-soft); color: var(--footwear); }}
        .embedded-sellers.apparel-seller .embedded-sellers-title {{ background: var(--apparel-soft); color: var(--apparel); }}
        .embedded-sellers.accessories-seller .embedded-sellers-title {{ background: var(--accessories-soft); color: var(--accessories); }}

        .embedded-sellers.footwear-seller .seller-title-icon {{ color: var(--footwear); stroke: var(--footwear); }}
        .embedded-sellers.apparel-seller .seller-title-icon {{ color: var(--apparel); stroke: var(--apparel); }}
        .embedded-sellers.accessories-seller .seller-title-icon {{ color: var(--accessories); stroke: var(--accessories); }}

        .embedded-sellers .data-table {{
            border: 0;
            border-radius: 0;
        }}

        .empty-note {{
            padding: 4mm;
            color: var(--muted);
            text-align: center;
            font-style: italic;
            background: var(--soft);
            border-radius: 7px;
        }}

        .footer {{
            margin-top: 4mm;
            padding-top: 2mm;
            border-top: 1px solid var(--line);
            color: #8795a5;
            text-align: center;
            font-size: 7.5px;
        }}

        @page {{
            size: Letter landscape;
            margin: 8mm 8mm 10mm 8mm;

            @bottom-left {{
                content: "Store Performance Reporting Platform";
                color: #7b8796;
                font-size: 7px;
            }}

            @bottom-right {{
                content: "Page " counter(page) " of " counter(pages);
                color: #7b8796;
                font-size: 7px;
            }}
        }}

        @page cover {{
            size: Letter landscape;
            margin: 0;

            @bottom-left {{ content: none; }}
            @bottom-right {{ content: none; }}
        }}


        /* ==========================================================
           Interactive analytics experience — browser only
           ========================================================== */

        html {{
            scroll-behavior: smooth;
            --report-zoom: 1;
        }}

        body.interactive-report {{
            padding-top: 58px;
        }}

        .interactive-shell {{
            position: fixed;
            inset: 0 0 auto 0;
            z-index: 2000;
            padding: 6px 12px;
            background: transparent;
            pointer-events: none;
        }}

        .interactive-toolbar {{
            width: min(1180px, calc(100% - 16px));
            margin: 0 auto;
            display: grid;
            grid-template-columns: auto minmax(0, 1fr) auto;
            align-items: center;
            gap: 8px;
            padding: 5px 7px;
            background: rgba(255, 255, 255, .78);
            border: 1px solid rgba(255, 255, 255, .72);
            border-radius: 11px;
            box-shadow: 0 3px 12px rgba(15, 39, 71, .045);
            backdrop-filter: blur(12px) saturate(115%);
            -webkit-backdrop-filter: blur(12px) saturate(115%);
            opacity: .62;
            pointer-events: auto;
            transition: opacity .18s ease, background .18s ease, box-shadow .18s ease;
        }}

        .interactive-toolbar:hover,
        .interactive-toolbar:focus-within {{
            opacity: 1;
            background: rgba(255, 255, 255, .88);
            box-shadow: 0 5px 18px rgba(15, 39, 71, .08);
        }}

        .interactive-brand {{
            display: flex;
            align-items: center;
            gap: 8px;
            color: var(--navy);
            font-weight: 800;
            white-space: nowrap;
        }}

        .live-dot {{
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: var(--green);
            box-shadow: 0 0 0 4px var(--green-soft);
        }}

        .interactive-tabs,
        .interactive-actions {{
            display: flex;
            align-items: center;
            justify-content: center;
            flex-wrap: wrap;
            gap: 6px;
        }}

        .interactive-actions {{
            justify-content: flex-end;
        }}

        .interactive-button,
        .department-filter {{
            min-height: 29px;
            padding: 5px 9px;
            border: 1px solid rgba(148, 163, 184, .28);
            border-radius: 8px;
            background: rgba(255, 255, 255, .48);
            color: var(--navy);
            font: inherit;
            font-size: 9px;
            font-weight: 700;
            cursor: pointer;
            transition: transform .15s ease, border-color .15s ease,
                        background .15s ease, box-shadow .15s ease;
        }}

        .interactive-button:hover,
        .department-filter:hover {{
            transform: translateY(-1px);
            border-color: #aebfd1;
            box-shadow: 0 4px 10px rgba(15, 39, 71, .08);
        }}

        .interactive-button:focus-visible,
        .department-filter:focus-visible,
        .report-search:focus-visible,
        .section-toggle:focus-visible {{
            outline: 2px solid var(--blue);
            outline-offset: 2px;
        }}

        .department-filter[aria-pressed="true"] {{
            color: var(--navy);
            border-color: rgba(15, 39, 71, .22);
            background: rgba(255, 255, 255, .92);
            box-shadow: 0 3px 9px rgba(15, 39, 71, .09);
        }}

        .department-filter[data-filter="footwear"][aria-pressed="true"] {{
            color: var(--footwear);
            border-color: rgba(23, 111, 166, .30);
            background: rgba(238, 245, 251, .92);
        }}

        .department-filter[data-filter="apparel"][aria-pressed="true"] {{
            color: var(--apparel);
            border-color: rgba(217, 106, 43, .30);
            background: rgba(255, 243, 234, .92);
        }}

        .department-filter[data-filter="accessories"][aria-pressed="true"] {{
            color: var(--accessories);
            border-color: rgba(74, 148, 85, .30);
            background: rgba(237, 248, 242, .92);
        }}

        .interactive-button.primary {{
            color: #1d4ed8;
            border-color: rgba(37, 99, 235, .24);
            background: rgba(234, 242, 255, .72);
        }}

        .text-size-controls {{
            display: inline-flex;
            overflow: hidden;
            border: 1px solid rgba(148, 163, 184, .28);
            border-radius: 8px;
            background: rgba(255, 255, 255, .42);
        }}

        .text-size-button {{
            min-width: 29px;
            min-height: 29px;
            padding: 4px 7px;
            border: 0;
            border-right: 1px solid rgba(148, 163, 184, .22);
            background: transparent;
            color: var(--navy);
            font: inherit;
            font-weight: 800;
            cursor: pointer;
        }}

        .text-size-button:last-child {{ border-right: 0; }}
        .text-size-button:hover {{ background: rgba(255, 255, 255, .72); }}

        /* Scale the complete browser report: text, cards, tables and charts. */
        body.interactive-report .page {{
            zoom: var(--report-zoom);
        }}

        .interaction-panel {{
            width: min(1180px, calc(100% - 24px));
            margin: -2px auto 14px;
            display: grid;
            grid-template-columns: minmax(240px, 1fr) auto;
            align-items: center;
            gap: 10px;
            padding: 10px 12px;
            background: rgba(255, 255, 255, .95);
            border: 1px solid rgba(148, 163, 184, .30);
            border-top: 0;
            border-radius: 0 0 12px 12px;
        }}

        .search-group {{
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .report-search {{
            width: min(460px, 100%);
            min-height: 34px;
            padding: 7px 10px;
            border: 1px solid #d6e0ea;
            border-radius: 8px;
            background: #fff;
            color: var(--ink);
            font: inherit;
            font-size: 10px;
        }}

        .interaction-status {{
            color: var(--muted);
            font-size: 8.5px;
            white-space: nowrap;
        }}

        .active-view-message {{
            display: flex;
            align-items: center;
            gap: 7px;
            color: var(--muted);
            font-size: 9px;
            text-align: right;
        }}

        .active-view-badge {{
            display: inline-flex;
            align-items: center;
            min-height: 24px;
            padding: 4px 8px;
            border-radius: 99px;
            background: var(--blue-soft);
            color: var(--blue);
            font-weight: 800;
        }}

        .screen-section {{
            scroll-margin-top: 118px;
        }}

        .section-title {{
            position: relative;
            padding-right: 30px;
        }}

        .section-toggle {{
            position: absolute;
            top: 50%;
            right: 0;
            width: 24px;
            height: 24px;
            transform: translateY(-50%);
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border: 1px solid #d6e0ea;
            border-radius: 7px;
            background: #fff;
            color: var(--navy);
            font-size: 13px;
            line-height: 1;
            cursor: pointer;
        }}

        .section.is-collapsed > :not(.section-title) {{
            display: none !important;
        }}

        .section.is-collapsed .section-toggle {{
            transform: translateY(-50%) rotate(-90deg);
        }}

        [data-department].is-filtered-out,
        .department-chart.is-filtered-out,
        .filterable-report-row.is-filtered-out {{
            display: none !important;
        }}

        .department-grid.filtered-view {{
            grid-template-columns: 1fr;
        }}

        .department-grid.filtered-view .department-card {{
            animation: reportReveal .22s ease-out;
        }}

        @keyframes reportReveal {{
            from {{ opacity: 0; transform: translateY(5px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .table-scroll {{
            width: 100%;
            overflow-x: auto;
            border-radius: 8px;
            -webkit-overflow-scrolling: touch;
        }}

        table.interactive-table th[data-sortable="true"] {{
            position: relative;
            padding-right: 4mm;
            cursor: pointer;
            user-select: none;
        }}

        table.interactive-table th[data-sortable="true"]::after {{
            content: "↕";
            position: absolute;
            right: 1.2mm;
            color: #94a3b8;
            font-size: 7px;
        }}

        table.interactive-table th[data-sort-direction="asc"]::after {{
            content: "↑";
            color: var(--blue);
        }}

        table.interactive-table th[data-sort-direction="desc"]::after {{
            content: "↓";
            color: var(--blue);
        }}

        table.interactive-table tbody tr {{
            transition: background .14s ease, opacity .14s ease;
        }}

        table.interactive-table tbody tr:hover td {{
            background: #eef5fb !important;
        }}

        .interactive-help {{
            position: fixed;
            right: 18px;
            bottom: 18px;
            z-index: 1900;
            display: flex;
            flex-direction: column;
            align-items: flex-end;
            gap: 7px;
        }}

        .help-card {{
            display: none;
            width: min(300px, calc(100vw - 36px));
            padding: 12px;
            border: 1px solid #d6e0ea;
            border-radius: 11px;
            background: rgba(255,255,255,.97);
            box-shadow: 0 10px 28px rgba(15,39,71,.14);
            color: var(--muted);
            font-size: 9px;
            line-height: 1.5;
        }}

        .help-card.is-open {{
            display: block;
        }}

        .help-card strong {{
            display: block;
            margin-bottom: 4px;
            color: var(--navy);
            font-size: 10px;
        }}

        .help-trigger,
        .back-to-top {{
            width: 40px;
            height: 40px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border: 1px solid #d6e0ea;
            border-radius: 50%;
            background: #fff;
            color: var(--navy);
            box-shadow: 0 7px 18px rgba(15,39,71,.13);
            font: inherit;
            font-size: 14px;
            font-weight: 800;
            cursor: pointer;
        }}

        .back-to-top {{
            display: none;
        }}

        .back-to-top.is-visible {{
            display: inline-flex;
        }}

        @media screen and (max-width: 1100px) {{
            body.interactive-report {{
                padding-top: 58px;
            }}

            .interactive-toolbar {{
                grid-template-columns: auto minmax(0, 1fr) auto;
                gap: 6px;
            }}

            .page {{
                width: calc(100% - 24px);
                min-height: auto;
                padding: 22px;
            }}

            .cover-page {{
                min-height: 640px;
            }}

            .cards {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}

            .metric-row {{
                grid-template-columns: repeat(3, minmax(0, 1fr));
            }}

            .brand-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
        }}

        @media screen and (max-width: 900px) and (min-width: 721px) {{
            .interactive-brand {{ display: none; }}
        }}

        @media screen and (max-width: 720px) {{
            html, body {{
                font-size: 13px;
            }}

            body.interactive-report {{
                padding-top: 0;
            }}

            .interactive-shell {{
                position: static;
                padding: 8px;
            }}

            .interactive-toolbar {{
                width: 100%;
                grid-template-columns: 1fr;
                padding: 7px;
            }}

            .interactive-tabs,
            .interactive-actions {{
                width: 100%;
            }}

            .department-filter,
            .interactive-button {{
                flex: 1 1 auto;
                min-height: 36px;
                font-size: 10px;
            }}

            .interaction-panel {{
                width: calc(100% - 16px);
                grid-template-columns: 1fr;
                margin-top: 0;
            }}

            .search-group {{
                align-items: stretch;
                flex-direction: column;
            }}

            .report-search {{
                width: 100%;
                min-height: 41px;
                font-size: 12px;
            }}

            .active-view-message {{
                justify-content: space-between;
                text-align: left;
            }}

            .page {{
                width: calc(100% - 16px);
                margin-bottom: 10px;
                padding: 14px;
                border-radius: 10px;
                overflow: visible;
            }}

            .cover-page {{
                min-height: auto;
            }}

            .cover-content {{
                padding: 22px 18px;
            }}

            .cover-main {{
                display: block;
                min-height: auto;
            }}

            .cover-panels {{
                display: none;
            }}

            .cover-meta-grid,
            .chart-row,
            .table-grid,
            .store-insights-grid,
            .seller-grid,
            .brand-grid {{
                grid-template-columns: 1fr;
            }}

            .cards {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 8px;
            }}

            .card {{
                min-height: 0;
                padding: 12px;
            }}

            .cover-bottom,
            .header {{
                align-items: flex-start;
                flex-direction: column;
            }}

            .header-meta {{
                min-width: 0;
                text-align: left;
            }}

            .single-chart-row .chart-card {{
                max-width: none;
            }}

            .metric-row {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}

            .data-table,
            .mini-table {{
                min-width: 700px;
            }}

            .interactive-help {{
                right: 10px;
                bottom: 10px;
            }}
        }}

        @media screen and (max-width: 430px) {{
            .interactive-brand {{ display: none; }}
            .interactive-tabs {{ gap: 4px; }}
            .department-filter {{
                flex: 1 1 calc(50% - 4px);
                padding-inline: 7px;
            }}
            .interactive-actions {{ gap: 4px; }}
            .interactive-button {{ padding-inline: 7px; }}
            .page {{ padding: 10px; }}
        }}

        /* ==========================================================
           Collapsible navigation rail
           ========================================================== */
        body.interactive-report {{
            padding-top: 0;
            padding-left: 176px;
            transition: padding-left .2s ease;
        }}

        body.interactive-report.nav-collapsed {{
            padding-left: 48px;
        }}

        .interactive-shell {{
            inset: 0 auto 0 0;
            width: 176px;
            padding: 10px 8px;
            display: flex;
            flex-direction: column;
            align-items: stretch;
            gap: 8px;
            pointer-events: none;
            transition: width .2s ease;
        }}

        .nav-toggle {{
            width: 36px;
            height: 36px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            align-self: flex-start;
            border: 1px solid rgba(148, 163, 184, .28);
            border-radius: 10px;
            background: rgba(255, 255, 255, .58);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            color: var(--navy);
            box-shadow: 0 3px 12px rgba(15, 39, 71, .06);
            font: inherit;
            font-size: 15px;
            cursor: pointer;
            pointer-events: auto;
        }}

        .interactive-toolbar {{
            width: 100%;
            margin: 0;
            grid-template-columns: 1fr;
            align-items: stretch;
            padding: 8px;
        }}

        .interactive-brand {{
            justify-content: flex-start;
            padding: 3px 4px 7px;
        }}

        .interactive-tabs,
        .interactive-actions {{
            width: 100%;
            flex-direction: column;
            align-items: stretch;
            gap: 5px;
        }}

        .department-filter,
        .interactive-button {{
            width: 100%;
            text-align: left;
        }}

        .text-size-controls {{
            width: 100%;
            margin-top: 3px;
        }}

        .text-size-button {{ flex: 1; }}

        .interactive-shell.is-collapsed {{ width: 48px; }}
        .interactive-shell.is-collapsed .interactive-toolbar {{ display: none; }}

        .screen-section {{ scroll-margin-top: 14px; }}

        @media screen and (max-width: 720px) {{
            body.interactive-report,
            body.interactive-report.nav-collapsed {{
                padding-left: 0;
                padding-top: 0;
            }}

            .interactive-shell,
            .interactive-shell.is-collapsed {{
                position: fixed;
                inset: 8px auto auto 8px;
                width: 42px;
                height: auto;
                padding: 0;
                z-index: 3000;
            }}

            .interactive-shell.is-open {{
                width: min(250px, calc(100vw - 16px));
            }}

            .nav-toggle {{
                width: 40px;
                height: 40px;
                background: rgba(255, 255, 255, .82);
            }}

            .interactive-shell .interactive-toolbar {{ display: none; }}
            .interactive-shell.is-open .interactive-toolbar {{
                display: grid;
                width: 100%;
                margin-top: 2px;
                opacity: 1;
                background: rgba(255, 255, 255, .94);
            }}

            .interactive-shell.is-open .interactive-brand {{ display: flex; }}
            .interactive-shell.is-open .interactive-tabs,
            .interactive-shell.is-open .interactive-actions {{
                flex-direction: column;
            }}

            .interactive-shell.is-open .department-filter,
            .interactive-shell.is-open .interactive-button {{
                width: 100%;
                min-height: 38px;
                text-align: left;
            }}
        }}

        /* Navigation stays completely out of the report until requested. */
        body.interactive-report,
        body.interactive-report.nav-collapsed {{
            padding-left: 0;
        }}

        .interactive-shell,
        .interactive-shell.is-collapsed {{
            position: fixed;
            inset: auto 18px 112px auto;
            width: 42px;
            height: auto;
            padding: 0;
            z-index: 3000;
            overflow: visible;
        }}

        .interactive-shell .interactive-toolbar,
        .interactive-shell.is-collapsed .interactive-toolbar {{
            display: none;
        }}

        .interactive-shell.is-open {{
            width: 42px;
        }}

        .interactive-shell.is-open .interactive-toolbar {{
            position: absolute;
            right: 0;
            bottom: 48px;
            display: grid;
            width: 220px;
            max-height: calc(100vh - 150px);
            overflow-y: auto;
            opacity: 1;
            background: rgba(255, 255, 255, .95);
            box-shadow: 0 10px 30px rgba(15, 39, 71, .16);
        }}

        .nav-toggle {{
            width: 42px;
            height: 42px;
            border-radius: 50%;
            background: rgba(255, 255, 255, .88);
        }}

        @media screen and (max-width: 720px) {{
            .interactive-shell,
            .interactive-shell.is-collapsed,
            .interactive-shell.is-open {{
                inset: auto 10px 108px auto;
                width: 42px;
            }}

            .interactive-shell.is-open .interactive-toolbar {{
                right: 0;
                bottom: 48px;
                width: min(230px, calc(100vw - 24px));
            }}

            /* Keep wide analytical tables inside the report canvas even when
               JavaScript is unavailable in a mobile document preview. */
            .section,
            .table-grid,
            .seller-block,
            .embedded-sellers {{
                min-width: 0;
                max-width: 100%;
                overflow-x: auto;
                -webkit-overflow-scrolling: touch;
            }}

            .data-table,
            .mini-table {{
                width: 100%;
                min-width: 620px;
                font-size: 8px;
            }}

            .product-table {{ min-width: 560px; }}

            .data-table th,
            .data-table td,
            .mini-table th,
            .mini-table td {{
                padding: 7px 6px;
            }}

            .page {{
                width: calc(100% - 16px);
                max-width: calc(100% - 16px);
                overflow-x: hidden;
            }}
        }}

        @media (prefers-reduced-motion: reduce) {{
            html {{
                scroll-behavior: auto;
            }}

            *,
            *::before,
            *::after {{
                animation: none !important;
                transition: none !important;
            }}
        }}

        @media print {{

            .interactive-shell,
            .interaction-panel,
            .interactive-help,
            .section-toggle {{
                display: none !important;
            }}

            body.interactive-report {{
                padding-top: 0 !important;
            }}

            body.interactive-report .page {{
                zoom: 1 !important;
            }}

            .section.is-collapsed > :not(.section-title),
            [data-department].is-filtered-out,
            .department-chart.is-filtered-out,
            .filterable-report-row.is-filtered-out {{
                display: block !important;
            }}

            tr.filterable-report-row.is-filtered-out {{
                display: table-row !important;
            }}

            .table-scroll {{
                overflow: visible !important;
            }}

            * {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}

            html, body {{
                background: #fff !important;
                width: auto !important;
                margin: 0 !important;
                padding: 0 !important;
            }}

            .page {{
                width: auto !important;
                min-height: auto !important;
                margin: 0 !important;
                padding: 0 !important;
                border-radius: 0 !important;
                box-shadow: none !important;
                overflow: visible !important;
                break-after: page;
                page-break-after: always;
            }}

            .cover-page {{
                width: 279.4mm !important;
                height: 215.9mm !important;
                min-height: 215.9mm !important;
                display: flex !important;
                overflow: hidden !important;
            }}

            .page:last-child {{
                break-after: auto;
                page-break-after: auto;
            }}

            .page-break {{
                break-before: auto !important;
                page-break-before: auto !important;
            }}

            .card, .chart-card, .section, .department-card,
            .brand-column, .seller-block, .store-insight-panel,
            .note-box, table, tr, img {{
                break-inside: avoid !important;
                page-break-inside: avoid !important;
                box-shadow: none !important;
            }}

            .seller-department-section {{
                break-inside: avoid !important;
                page-break-inside: avoid !important;
            }}

            .department-grid {{
                display: block !important;
            }}

            .department-grid .department-card {{
                margin-bottom: 4mm;
                /* Department cards can exceed one physical PDF page. Keeping
                   the whole card unbroken pushes it past the page header and
                   leaves a nearly blank page behind. */
                break-inside: auto !important;
                page-break-inside: auto !important;
                /* overflow:hidden makes WeasyPrint treat the card like one
                   non-fragmentable box even when break-inside is overridden. */
                overflow: visible !important;
            }}

            .page > .header {{
                break-after: avoid !important;
                page-break-after: avoid !important;
            }}

            /* Keep each internal unit polished even when the outer department
               card continues onto the following page. */
            .department-header,
            .metric-row,
            .department-subsection-heading {{
                break-inside: avoid !important;
                page-break-inside: avoid !important;
                break-after: avoid !important;
                page-break-after: avoid !important;
            }}

            .brand-grid {{
                break-inside: auto !important;
                page-break-inside: auto !important;
            }}

            thead {{
                display: table-header-group;
            }}

            tfoot {{
                display: table-footer-group;
            }}

            img {{
                max-width: 100% !important;
                height: auto !important;
            }}

            .chart-card img {{
                max-height: 48mm !important;
            }}

            .top-brands-chart img {{
                max-height: 53mm !important;
            }}

            .header {{
                margin-top: 0 !important;
            }}

            .section {{
                margin-top: 3mm !important;
            }}

            .footer {{
                margin-top: 3mm !important;
            }}
        }}
    </style>

</head>
<body class="interactive-report">

    <div class="interactive-shell" id="navigationRail" role="region" aria-label="Interactive report controls">
        <button class="nav-toggle" id="navigationToggle" type="button" aria-expanded="true" aria-controls="reportToolbar" aria-label="Collapse report navigation">☰</button>
        <div class="interactive-toolbar" id="reportToolbar">
            <div class="interactive-brand">
                <span class="live-dot" aria-hidden="true"></span>
                <span>Interactive Report</span>
            </div>

            <div class="interactive-tabs" role="group" aria-label="Department view">
                <button class="department-filter" type="button" data-filter="all" data-jump="overview" aria-pressed="true">Overview</button>
                <button class="department-filter" type="button" data-filter="footwear" aria-pressed="false">Footwear</button>
                <button class="department-filter" type="button" data-filter="apparel" aria-pressed="false">Apparel</button>
                <button class="department-filter" type="button" data-filter="accessories" aria-pressed="false">Accessories</button>
                <button class="interactive-button" type="button" data-jump="departments">Departments</button>
                <button class="interactive-button" type="button" data-jump="opportunities">Opportunities</button>
            </div>

            <div class="interactive-actions">
                <div class="text-size-controls" role="group" aria-label="Report zoom">
                    <button class="text-size-button" id="decreaseText" type="button" aria-label="Zoom report out">A−</button>
                    <button class="text-size-button" id="increaseText" type="button" aria-label="Zoom report in">A+</button>
                </div>
                <button class="interactive-button primary" id="printReport" type="button">Print / PDF</button>
            </div>
        </div>
    </div>


    <div class="page cover-page">
        <div class="cover-accent"></div>
        <div class="cover-content">
            <div class="cover-topbar"><div class="cover-eyebrow">Store Performance Reporting Platform</div>{cover_logo}</div>
            <div class="cover-main">
                <div class="cover-left">
                    <div class="cover-kicker">Operational Intelligence</div>
                    <h1 class="cover-title">{report_title}</h1>
                    <div class="cover-subtitle">Executive KPI Dashboard</div>
                    <div class="cover-store">{store_name}</div>
                    <div class="cover-rule"><span class="rule-footwear"></span><span class="rule-apparel"></span><span class="rule-accessories"></span></div>

                    <div class="cover-meta-grid">
                        <div class="cover-meta-item"><span>Reporting period</span><strong>{report_period}</strong></div>
                        <div class="cover-meta-item"><span>Generated</span><strong>{generated_date}</strong></div>
                        <div class="cover-meta-item"><span>Report scope</span><strong>Sales, budget, brands &amp; inventory</strong></div>
                        <div class="cover-meta-item"><span>Source</span><strong>Standardized Tableau exports</strong></div>
                    </div>
                </div>

                <div class="cover-panels">
                    <div class="cover-dept-panel footwear-panel">
                        <span class="cover-dept-icon-shell">{cover_section_icon('Footwear')}</span><div class="cover-dept-name">Footwear</div><div class="cover-mini-rule"></div>
                        <div class="cover-momentum"><span></span><span></span><span></span><span></span><span></span></div>
                        <div class="cover-mini-bars"><span style="height:24%"></span><span style="height:38%"></span><span style="height:31%"></span><span style="height:55%"></span><span style="height:70%"></span></div>
                    </div>
                    <div class="cover-dept-panel apparel-panel">
                        <span class="cover-dept-icon-shell">{cover_section_icon('Apparel')}</span><div class="cover-dept-name">Apparel</div><div class="cover-mini-rule"></div>
                        <div class="cover-momentum"><span></span><span></span><span></span><span></span><span></span></div>
                        <div class="cover-mini-bars"><span style="height:22%"></span><span style="height:35%"></span><span style="height:48%"></span><span style="height:66%"></span><span style="height:82%"></span></div>
                    </div>
                    <div class="cover-dept-panel accessories-panel">
                        <span class="cover-dept-icon-shell">{cover_section_icon('Accessories')}</span><div class="cover-dept-name">Accessories</div><div class="cover-mini-rule"></div>
                        <div class="cover-momentum"><span></span><span></span><span></span><span></span><span></span></div>
                        <div class="cover-mini-bars"><span style="height:20%"></span><span style="height:32%"></span><span style="height:46%"></span><span style="height:65%"></span><span style="height:88%"></span></div>
                    </div>
                </div>
            </div>

            <div class="cover-bottom">
                <div class="cover-description">
                    Automated operational report consolidating store KPIs,
                    department performance, sales contribution, top brands,
                    top-selling products, budget opportunities and validation checks.
                    <div class="cover-contents">Executive KPIs &nbsp;·&nbsp; Department Performance &nbsp;·&nbsp; Top Sellers &nbsp;·&nbsp; Budget Opportunities</div>
                    <div class="cover-credit">Report design and automation: Diego Diaz Iturbe</div>
                </div>
                <div class="cover-version">Report version 1.0</div>
            </div>
        </div>
    </div>

    <div class="page page-break screen-section" id="overview">
        <div class="header">
            <div><div class="title">{store_name} KPI Report</div><div class="subtitle">{report_title}</div></div>
            <div class="header-actions">
                <div class="header-meta"><div><strong>Period:</strong> {report_period}</div><div><strong>Generated:</strong> {generated_date}</div></div>
            </div>
        </div>

        <div class="cards">
            {kpi_card("Net Sales", money(net_sales), f"{pct(sales_vs_ly)} vs LY", "sales", budget_attainment, f"{pct(budget_attainment)} Budget attainment", "positive" if budget_attainment >= 1 else "negative")}
            {kpi_card("Budget", money(budget), "Reporting-period target", "budget", None, "", "neutral")}
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
            <div class="chart-card compact-chart"><h3>Budget Gap by Department</h3><img src="../assets/charts/budget_gap.png" alt="Budget Gap by Department"></div>
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

    <div class="page page-break screen-section" id="departments">
        <div class="header">
            <div><div class="title">Department Detail</div><div class="subtitle">Footwear, Apparel and Accessories Performance</div></div>
            <div class="header-meta"><div><strong>Store:</strong> {store_name}</div><div><strong>Period:</strong> {report_period}</div></div>
        </div>
        <div class="department-grid">{department_cards}</div>


        <div class="section">
            <div class="section-title">Subdepartment Contribution</div>

            <div class="chart-row">
                <div class="chart-card compact-chart department-chart" data-department="footwear"><h3>{section_icon("Footwear", "chart-heading-icon")}<span>Footwear Contribution</span></h3><img src="../assets/charts/footwear_mix.png" alt="Footwear Contribution"></div>
                <div class="chart-card compact-chart department-chart" data-department="apparel"><h3>{section_icon("Apparel", "chart-heading-icon")}<span>Apparel Contribution</span></h3><img src="../assets/charts/apparel_mix.png" alt="Apparel Contribution"></div>
            </div>

            <div class="chart-row single-chart-row">
                <div class="chart-card compact-chart department-chart" data-department="accessories"><h3>{section_icon("Accessories", "chart-heading-icon")}<span>Accessories Contribution</span></h3><img src="../assets/charts/accessories_mix.png" alt="Accessories Contribution"></div>
            </div>
        </div>
        <div class="footer">Department sections generated automatically from Tableau brand exports.</div>
    </div>

    <div class="page page-break screen-section" id="opportunities">
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

    <div class="interactive-help">
        <div class="help-card" id="helpCard">
            <strong>How to interact</strong>
            Choose a department to focus the detail pages, search across every
            table, click a table heading to sort it, collapse sections using the
            arrow beside their title, or jump directly to a report area.
        </div>
        <button class="help-trigger" id="helpTrigger" type="button" aria-expanded="false" aria-controls="helpCard" aria-label="Show interaction help">?</button>
        <button class="back-to-top" id="backToTop" type="button" aria-label="Back to overview">↑</button>
    </div>

    <script>
    (() => {{
        "use strict";

        const departmentNames = {{
            all: "Overall report",
            footwear: "Footwear",
            apparel: "Apparel",
            accessories: "Accessories"
        }};

        const normalize = (value) =>
            String(value ?? "")
                .trim()
                .toLocaleLowerCase()
                .replace(/\\s+/g, " ");

        const parseNumber = (value) => {{
            const cleaned = String(value ?? "")
                .replace(/[$,%]/g, "")
                .replace(/,/g, "")
                .replace(/[^\\d.+-]/g, "");
            const parsed = Number(cleaned);
            return Number.isFinite(parsed) ? parsed : null;
        }};

        const filterButtons = Array.from(document.querySelectorAll(".department-filter"));
        const departmentElements = Array.from(document.querySelectorAll("[data-department]"))
            .filter((element) => !element.classList.contains("department-filter"));
        const departmentGrid = document.querySelector(".department-grid");
        const activeViewBadge = document.getElementById("activeViewBadge");
        const searchInput = document.getElementById("reportSearch");
        const searchStatus = document.getElementById("searchStatus");
        const tables = Array.from(document.querySelectorAll("table"));
        const sections = Array.from(document.querySelectorAll(".section"));
        const backToTop = document.getElementById("backToTop");
        const navigationRail = document.getElementById("navigationRail");
        const navigationToggle = document.getElementById("navigationToggle");
        const mobileNavigation = window.matchMedia("(max-width: 720px)");
        let activeDepartment = "all";

        const setNavigationState = (open) => {{
            navigationRail?.classList.toggle("is-open", open);
            navigationRail?.classList.toggle("is-collapsed", !open);
            document.body.classList.remove("nav-collapsed");
            navigationToggle?.setAttribute("aria-expanded", String(open));
            navigationToggle?.setAttribute(
                "aria-label",
                `${{open ? "Collapse" : "Expand"}} report navigation`
            );
        }};

        navigationToggle?.addEventListener("click", () => {{
            const currentlyOpen = navigationRail?.classList.contains("is-open");
            setNavigationState(!currentlyOpen);
        }});

        mobileNavigation.addEventListener?.("change", () => setNavigationState(false));
        setNavigationState(false);

        // Tag table rows with department context where their content makes it clear.
        const departmentFromText = (text) => {{
            const value = normalize(text);
            if (value.includes("footwear")) return "footwear";
            if (value.includes("apparel")) return "apparel";
            if (value.includes("accessories")) return "accessories";
            return "";
        }};

        tables.forEach((table) => {{
            table.classList.add("interactive-table");

            if (!table.parentElement?.classList.contains("table-scroll")) {{
                const wrapper = document.createElement("div");
                wrapper.className = "table-scroll";
                table.parentNode.insertBefore(wrapper, table);
                wrapper.appendChild(table);
            }}

            Array.from(table.tBodies).forEach((body) => {{
                Array.from(body.rows).forEach((row) => {{
                    row.classList.add("filterable-report-row");
                    const inferredDepartment = departmentFromText(row.textContent);
                    if (inferredDepartment) row.dataset.rowDepartment = inferredDepartment;
                }});
            }});

            const body = table.tBodies[0];
            if (!body) return;

            const headers = Array.from(table.querySelectorAll("thead th"));
            headers.forEach((header, columnIndex) => {{
                header.dataset.sortable = "true";
                header.setAttribute("role", "button");
                header.tabIndex = 0;
                header.setAttribute("aria-label", `Sort by ${{header.textContent.trim()}}`);

                const sort = () => {{
                    const direction = header.dataset.sortDirection === "asc" ? "desc" : "asc";
                    headers.forEach((item) => delete item.dataset.sortDirection);
                    header.dataset.sortDirection = direction;

                    const rows = Array.from(body.rows);
                    rows.sort((rowA, rowB) => {{
                        const rawA = rowA.cells[columnIndex]?.textContent ?? "";
                        const rawB = rowB.cells[columnIndex]?.textContent ?? "";
                        const numA = parseNumber(rawA);
                        const numB = parseNumber(rawB);
                        const comparison =
                            numA !== null && numB !== null
                                ? numA - numB
                                : rawA.localeCompare(rawB, undefined, {{
                                    numeric: true,
                                    sensitivity: "base"
                                }});
                        return direction === "asc" ? comparison : -comparison;
                    }});

                    rows.forEach((row) => body.appendChild(row));
                }};

                header.addEventListener("click", sort);
                header.addEventListener("keydown", (event) => {{
                    if (event.key === "Enter" || event.key === " ") {{
                        event.preventDefault();
                        sort();
                    }}
                }});
            }});
        }});

        const applyView = () => {{
            const query = normalize(searchInput?.value);
            let visibleRows = 0;
            let totalRows = 0;

            departmentElements.forEach((element) => {{
                const matchesDepartment =
                    activeDepartment === "all" ||
                    element.dataset.department === activeDepartment;
                element.classList.toggle("is-filtered-out", !matchesDepartment);
            }});

            departmentGrid?.classList.toggle("filtered-view", activeDepartment !== "all");

            document.querySelectorAll(".filterable-report-row").forEach((row) => {{
                totalRows += 1;
                const matchesSearch = !query || normalize(row.textContent).includes(query);
                const rowDepartment = row.dataset.rowDepartment || "";
                const matchesDepartment =
                    activeDepartment === "all" ||
                    !rowDepartment ||
                    rowDepartment === activeDepartment;
                const visible = matchesSearch && matchesDepartment;
                row.classList.toggle("is-filtered-out", !visible);
                if (visible) visibleRows += 1;
            }});

            if (searchStatus) {{
                searchStatus.textContent =
                    query || activeDepartment !== "all"
                        ? `${{visibleRows}} of ${{totalRows}} rows shown`
                        : `${{totalRows}} searchable rows`;
            }}

            if (activeViewBadge) {{
                activeViewBadge.textContent = departmentNames[activeDepartment];
            }}
        }};

        filterButtons.forEach((button) => {{
            button.addEventListener("click", () => {{
                activeDepartment = button.dataset.filter || "all";
                filterButtons.forEach((item) =>
                    item.setAttribute("aria-pressed", String(item === button))
                );
                applyView();

                if (activeDepartment !== "all") {{
                    document.getElementById("departments")?.scrollIntoView({{
                        behavior: "smooth",
                        block: "start"
                    }});
                }}
                if (mobileNavigation.matches) setNavigationState(false);
            }});
        }});

        searchInput?.addEventListener("input", applyView);

        // Each report section can be collapsed independently.
        sections.forEach((section) => {{
            const title = section.querySelector(":scope > .section-title");
            if (!title) return;

            const button = document.createElement("button");
            button.className = "section-toggle";
            button.type = "button";
            button.innerHTML = "⌄";
            button.setAttribute("aria-expanded", "true");
            button.setAttribute("aria-label", `Collapse ${{title.textContent.trim()}}`);

            button.addEventListener("click", () => {{
                const collapsed = section.classList.toggle("is-collapsed");
                button.setAttribute("aria-expanded", String(!collapsed));
                button.setAttribute(
                    "aria-label",
                    `${{collapsed ? "Expand" : "Collapse"}} ${{title.textContent.trim()}}`
                );
            }});

            title.appendChild(button);
        }});

        document.querySelectorAll("[data-jump]").forEach((button) => {{
            button.addEventListener("click", () => {{
                document.getElementById(button.dataset.jump)?.scrollIntoView({{
                    behavior: "smooth",
                    block: "start"
                }});
                if (mobileNavigation.matches) setNavigationState(false);
            }});
        }});

        document.getElementById("printReport")?.addEventListener("click", () => {{
            sections.forEach((section) => section.classList.remove("is-collapsed"));
            window.print();
        }});

        const helpTrigger = document.getElementById("helpTrigger");
        const helpCard = document.getElementById("helpCard");
        helpTrigger?.addEventListener("click", () => {{
            const open = helpCard?.classList.toggle("is-open") ?? false;
            helpTrigger.setAttribute("aria-expanded", String(open));
        }});

        window.addEventListener("scroll", () => {{
            backToTop?.classList.toggle("is-visible", window.scrollY > 500);
        }}, {{ passive: true }});

        backToTop?.addEventListener("click", () => {{
            document.getElementById("overview")?.scrollIntoView({{
                behavior: "smooth",
                block: "start"
            }});
        }});

        // Discreet accessibility control for report text. The setting is kept
        // for the next time the same report is opened in this browser.
        const scaleKey = "ddi-report-zoom";
        const clampScale = (value) => Math.min(1.3, Math.max(.8, value));
        const readSavedScale = () => {{
            try {{ return Number(localStorage.getItem(scaleKey)) || 1; }}
            catch (_error) {{ return 1; }}
        }};
        let textScale = clampScale(readSavedScale());
        const applyTextScale = () => {{
            document.documentElement.style.setProperty("--report-zoom", textScale);
            try {{ localStorage.setItem(scaleKey, String(textScale)); }}
            catch (_error) {{ /* Local files may block storage; sizing still works. */ }}
        }};

        document.getElementById("decreaseText")?.addEventListener("click", () => {{
            textScale = clampScale(textScale - .1);
            applyTextScale();
        }});

        document.getElementById("increaseText")?.addEventListener("click", () => {{
            textScale = clampScale(textScale + .1);
            applyTextScale();
        }});

        window.addEventListener("beforeprint", () => {{
            sections.forEach((section) => section.classList.remove("is-collapsed"));
        }});

        applyTextScale();
        applyView();
    }})();
    </script>

</body>
</html>
"""

    output_path = OUTPUT_DIR / "final_report.html"
    output_path.write_text(html, encoding="utf-8")
    print(f"Created: {output_path}")
    return output_path


if __name__ == "__main__":
    build_html_report()
