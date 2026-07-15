import matplotlib.pyplot as plt
import pandas as pd
from pathlib import Path

CHART_DIR = Path("assets/charts")
CHART_DIR.mkdir(parents=True, exist_ok=True)

POSITIVE = "#43A047"
NEGATIVE = "#E53935"
BLUE = "#176FA6"
ORANGE = "#D96A2B"
GREEN = "#4A9455"
TEXT = "#2B2F33"
GRID = "#E2E8F0"

DEPARTMENT_COLORS = {
    "footwear": BLUE,
    "apparel": ORANGE,
    "accessories": GREEN,
}
DEPARTMENT_ORDER = {"footwear": 0, "apparel": 1, "accessories": 2}

# Dark-to-light, preserving one family per department.
GROUP_PALETTES = {
    "footwear": ["#155F8E", "#2879A8", "#3D8DB8", "#559FC4", "#78B4D2"],
    "apparel": ["#A94D20", "#C76028", "#D96A2B", "#E98B57", "#F0AD88"],
    "accessories": ["#34753F", "#4A9455", "#70AE79", "#91C497", "#B2D7B7"],
}


def money(value):
    return f"${float(value):,.0f}"


def _save(fig, output_name, pad=0.15):
    output_path = CHART_DIR / output_name
    fig.savefig(output_path, bbox_inches="tight", pad_inches=pad, facecolor="white")
    plt.close(fig)
    return output_path


def _autopct_visible(threshold=2.0):
    def formatter(pct):
        return f"{pct:.0f}%" if pct >= threshold else ""
    return formatter


def make_budget_gap_chart(df, output_name="budget_gap.png"):
    df = df.copy()
    df["Budget Gap"] = pd.to_numeric(df["Budget Gap"], errors="coerce").fillna(0)
    df = df[~df["Department"].astype(str).str.lower().str.contains("total")]
    df["_order"] = df["Department"].astype(str).str.lower().map(DEPARTMENT_ORDER).fillna(99)
    df = df.sort_values("_order", ascending=False)

    labels = df["Department"].astype(str).tolist()
    values = df["Budget Gap"].tolist()
    colors = [POSITIVE if value > 0 else NEGATIVE if value < 0 else "#94A3B8" for value in values]

    fig, ax = plt.subplots(figsize=(7.1, 2.45), dpi=300)
    bars = ax.barh(
        labels,
        values,
        color=colors,
        height=0.5,
        edgecolor="#334155",
        linewidth=0.35,
    )
    for bar, value in zip(bars, values):
        if value < 0:
            bar.set_hatch("////")
        elif value == 0:
            bar.set_hatch("..")

    ax.axvline(0, color="#0F172A", linewidth=1.0)
    max_abs = max([abs(v) for v in values] + [1])

    for bar, value in zip(bars, values):
        y = bar.get_y() + bar.get_height() / 2
        offset = max_abs * 0.025
        ax.text(
            value + offset if value >= 0 else value - offset,
            y,
            money(value) if value >= 0 else f"-{money(abs(value))}",
            va="center",
            ha="left" if value >= 0 else "right",
            fontsize=8.5,
            weight="bold",
            color=TEXT,
        )

    ax.set_xlim(-max_abs * 1.18, max_abs * 1.18)
    ax.grid(axis="x", color=GRID, linewidth=0.55)
    ax.set_axisbelow(True)
    ax.margins(y=0.12)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(axis="x", labelsize=7.0, colors="#5B6470", pad=2)
    ax.tick_params(axis="y", labelsize=9, colors=TEXT, length=0, pad=6)
    fig.subplots_adjust(left=0.18, right=0.93, top=0.94, bottom=0.18)
    return _save(fig, output_name)


def make_sales_mix_donut(df, output_name="sales_mix.png"):
    df = df.copy()
    df["Net Sales"] = pd.to_numeric(df["Net Sales"], errors="coerce").fillna(0)
    df["Sales %"] = pd.to_numeric(df["Sales %"], errors="coerce").fillna(0)
    df = df[~df["Department"].astype(str).str.lower().str.contains("total")]
    df["_order"] = df["Department"].astype(str).str.lower().map(DEPARTMENT_ORDER).fillna(99)
    df = df.sort_values("_order")

    labels = df["Department"].astype(str).tolist()
    values = df["Net Sales"].tolist()
    total = sum(values)
    colors = [DEPARTMENT_COLORS.get(label.lower(), "#64748B") for label in labels]

    fig, ax = plt.subplots(figsize=(5.7, 3.45), dpi=300)
    wedges, _, _ = ax.pie(
        values,
        labels=None,
        colors=colors,
        startangle=90,
        counterclock=False,
        wedgeprops=dict(width=0.40, edgecolor="white", linewidth=1.2),
        autopct=_autopct_visible(2.0),
        pctdistance=0.79,
        textprops=dict(color="white", fontsize=8.5, weight="bold"),
    )

    ax.text(0, 0.05, money(total), ha="center", va="center",
            fontsize=12.5, weight="bold", color=TEXT)
    ax.text(0, -0.12, "Total Sales", ha="center", va="center",
            fontsize=7.5, color="#667085")

    legend_labels = [
        f"{row['Department']}  {money(row['Net Sales'])} ({row['Sales %']:.0f}%)"
        for _, row in df.iterrows()
    ]
    ax.legend(
        wedges,
        legend_labels,
        loc="lower center",
        bbox_to_anchor=(0.5, -0.08),
        ncol=min(3, len(legend_labels)),
        frameon=False,
        fontsize=7.2,
        handlelength=1.4,
        columnspacing=1.2,
    )
    ax.set(aspect="equal")
    fig.subplots_adjust(left=0.03, right=0.97, top=0.98, bottom=0.20)
    return _save(fig, output_name)


def make_top_brands_chart(df, output_name="top_brands.png", title=None):
    df = df.copy()
    df["Brand"] = df["Brand"].astype(str).str.strip()
    df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce").fillna(0)

    group_col = "Group" if "Group" in df.columns else ("group" if "group" in df.columns else None)
    grouping = ["Brand"] + ([group_col] if group_col else [])
    df = (
        df.groupby(grouping, as_index=False)["Sales"]
        .sum()
        .sort_values("Sales", ascending=False)
        .head(8)
        .sort_values("Sales", ascending=True)
    )

    if group_col:
        df["Label"] = df["Brand"] + " — " + df[group_col].astype(str)
        colors = [
            DEPARTMENT_COLORS.get(str(group).lower(), "#64748B")
            for group in df[group_col]
        ]
    else:
        df["Label"] = df["Brand"]
        colors = [BLUE] * len(df)

    labels = df["Label"].tolist()
    values = df["Sales"].tolist()
    fig, ax = plt.subplots(figsize=(5.35, 3.2), dpi=300)
    bars = ax.barh(labels, values, color=colors, height=0.43)

    max_value = max(values) if values else 1
    for bar, value in zip(bars, values):
        ax.text(
            value + max_value * 0.018,
            bar.get_y() + bar.get_height() / 2,
            money(value),
            va="center",
            ha="left",
            fontsize=7.6,
            weight="bold",
            color=TEXT,
        )

    ax.set_xlim(0, max_value * 1.18)
    ax.grid(axis="x", color=GRID, linewidth=0.55)
    ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(axis="x", labelsize=7.5, colors="#5B6470", pad=2)
    ax.tick_params(axis="y", labelsize=7.7, colors=TEXT, length=0, pad=4)
    fig.subplots_adjust(left=0.40, right=0.90, top=0.97, bottom=0.16)
    return _save(fig, output_name)


def make_subdepartment_mix_chart(df, group_name, output_name):
    df = df.copy()
    df = df[df["group"].astype(str).str.lower() == group_name.lower()]
    df = df[~df["department"].astype(str).str.lower().str.contains("total")]
    df["net_sales"] = pd.to_numeric(df["net_sales"], errors="coerce").fillna(0)
    df = df[df["net_sales"] > 0].copy()

    order_map = {
        "mens": 0, "men's": 0,
        "womens": 1, "women's": 1,
        "junior": 2,
        "childrens": 3, "children's": 3,
        "infants": 4, "infant": 4,
        "clothing accessories": 0,
        "other accessories": 1,
    }

    def sub_order(value):
        text = str(value).lower()
        for key, rank in order_map.items():
            if key in text:
                return rank
        return 99

    df["_order"] = df["department"].apply(sub_order)
    df = df.sort_values("_order")

    labels = df["department"].astype(str).tolist()
    values = df["net_sales"].tolist()
    total = sum(values)
    palette = GROUP_PALETTES.get(group_name.lower(), ["#64748B"])
    colors = [palette[min(i, len(palette) - 1)] for i in range(len(values))]

    fig, ax = plt.subplots(figsize=(5.35, 3.45), dpi=300)
    wedges, _, _ = ax.pie(
        values,
        labels=None,
        colors=colors,
        startangle=90,
        counterclock=False,
        wedgeprops=dict(width=0.40, edgecolor="white", linewidth=1.2),
        autopct=_autopct_visible(2.0),
        pctdistance=0.79,
        textprops=dict(color="white", fontsize=8.2, weight="bold"),
    )

    ax.text(0, 0.05, money(total), ha="center", va="center",
            fontsize=12, weight="bold", color=TEXT)
    ax.text(0, -0.12, "Sales", ha="center", va="center",
            fontsize=7.5, color="#667085")

    legend_labels = [
        f"{row['department']}  {money(row['net_sales'])}"
        for _, row in df.iterrows()
    ]
    ax.legend(
        wedges,
        legend_labels,
        loc="lower center",
        bbox_to_anchor=(0.5, -0.10),
        ncol=2 if len(legend_labels) > 2 else len(legend_labels),
        frameon=False,
        fontsize=6.9,
        handlelength=1.3,
        columnspacing=1.0,
    )
    ax.set(aspect="equal")
    fig.subplots_adjust(left=0.02, right=0.98, top=0.98, bottom=0.22)
    return _save(fig, output_name)
