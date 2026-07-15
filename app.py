"""Store Performance Reporting Platform - Streamlit App.

Current workflow:
- Enter the store/location and report dates manually.
- Upload all eight Tableau exports in one bulk uploader.
- Identify and validate each file automatically.
- Keep the Generate Report button visible but disabled until inputs are complete.
- Build the data model, calculations, charts, HTML report, and optional PDF.
"""

from __future__ import annotations

import base64
import hashlib
import html
import tempfile
from datetime import date
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd
import streamlit as st

from upload_classifier import (
    REPORT_DEFINITIONS,
    classify_uploaded_files,
    validation_rows,
)
from vm_report_data_model_v2 import INPUT_FILES, build_model, export_model
from report_calculations import (
    build_report_calculations,
    export_calculations_to_excel,
    money,
    pct,
    number,
)
from html_report_builder import build_html_report


APP_DIR = Path(__file__).resolve().parent
ASSETS_DIR = APP_DIR / "assets"
HERO_IMAGE_PATH = ASSETS_DIR / "store_performance_hero.png"
FAVICON_PATH = ASSETS_DIR / "ddi_favicon.png"

st.set_page_config(
    page_title="Store Performance Reporting Platform",
    page_icon=str(FAVICON_PATH) if FAVICON_PATH.exists() else "📈",
    layout="wide",
)


def image_data_uri(image_path: Path) -> str | None:
    """Return a data URI for a local image, or None when the asset is absent."""
    if not image_path.exists():
        return None

    suffix = image_path.suffix.lower().lstrip(".")
    mime_subtype = "jpeg" if suffix in {"jpg", "jpeg"} else suffix
    encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
    return f"data:image/{mime_subtype};base64,{encoded}"


def render_platform_hero() -> None:
    """Render the product hero image with a graceful CSS fallback."""
    hero_uri = image_data_uri(HERO_IMAGE_PATH)

    if hero_uri:
        st.markdown(
            f"""
            <section
                class="vm-hero vm-hero-image"
                aria-label="Retail Performance Reporting Platform"
                style="background-image:url('{hero_uri}');"
            >
                <div class="vm-hero-overlay"></div>
                <div class="vm-hero-copy vm-hero-copy-overlay">
                    <div class="vm-kicker">DDI DATA SOLUTIONS</div>
                    <div class="vm-title">Retail Performance Reporting Platform</div>
                    <div class="vm-sub">
                        Transform Tableau exports into validated, executive-ready KPI reports
                        with automated analysis and HTML, Excel and PDF outputs.
                    </div>
                    <div class="vm-author">
                        Created by Diego Díaz Iturbe · Retail Analytics · Reporting Automation
                    </div>
                    <div class="vm-version">Cloud application · 8-source automated validation</div>
                </div>
            </section>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        """
        <section class="vm-hero vm-hero-fallback">
            <div class="vm-hero-copy">
                <div class="vm-kicker">DDI DATA SOLUTIONS</div>
                <div class="vm-title">Retail Performance Reporting Platform</div>
                <div class="vm-sub">
                    Transform Tableau exports into validated, executive-ready KPI reports with
                    automated analysis and HTML, Excel and PDF outputs.
                </div>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


st.markdown("""
<style>
.block-container{
    max-width:1500px;
    padding-top:.8rem;
    padding-bottom:3rem;
}
.vm-hero{
    position:relative;
    overflow:hidden;
    width:100%;
    min-height:355px;
    margin:0 0 1.15rem 0;
    border-radius:0 0 24px 24px;
    color:white;
    box-shadow:0 18px 40px rgba(15,23,42,.18);
}
.vm-hero-image{
    display:flex;
    align-items:center;
    background-size:cover;
    background-position:center center;
    background-repeat:no-repeat;
}
.vm-hero-overlay{
    position:absolute;
    inset:0;
    background:
        linear-gradient(90deg,rgba(3,12,30,.92) 0%,rgba(4,19,45,.76) 44%,rgba(4,18,40,.24) 76%,rgba(4,12,28,.12) 100%),
        linear-gradient(0deg,rgba(3,12,30,.32),transparent 55%);
}
.vm-hero-copy-overlay{
    position:relative;
    z-index:1;
    max-width:760px;
    padding:3.25rem 3rem 2.65rem;
    text-shadow:0 2px 16px rgba(0,0,0,.34);
}
.vm-author{
    display:inline-flex;
    margin-top:1.15rem;
    padding:.42rem .72rem;
    border:1px solid rgba(255,255,255,.42);
    border-radius:999px;
    background:rgba(3,12,30,.24);
    color:#f8fafc;
    font-size:.79rem;
    font-weight:650;
    backdrop-filter:blur(4px);
}
.vm-version{margin-top:.65rem;color:#cbd5e1;font-size:.72rem;font-weight:650;}
.vm-hero-fallback{
    display:flex;
    align-items:center;
    padding:2.5rem;
    background:
        radial-gradient(circle at 82% 20%,rgba(45,212,191,.24),transparent 32%),
        linear-gradient(135deg,#07152d,#163b77 58%,#0f766e);
}
.vm-hero-copy{max-width:760px;}
.vm-kicker{
    margin-bottom:.7rem;
    color:#bfdbfe;
    font-size:.78rem;
    font-weight:850;
    letter-spacing:.16em;
}
.vm-title{
    margin-bottom:.7rem;
    font-size:clamp(2rem,4vw,3.2rem);
    line-height:1.04;
    font-weight:900;
    letter-spacing:-.035em;
}
.vm-sub{
    max-width:760px;
    color:#e2e8f0;
    font-size:1.04rem;
    line-height:1.55;
}
.sr-only{
    position:absolute;
    width:1px;
    height:1px;
    padding:0;
    margin:-1px;
    overflow:hidden;
    clip:rect(0,0,0,0);
    white-space:nowrap;
    border:0;
}
@media(max-width:900px){
    .vm-hero{min-height:300px;border-radius:0 0 18px 18px;}
    .vm-hero-image{background-position:center center;}
}
@media(max-width:620px){
    .vm-hero{min-height:330px;}
    .vm-hero-image{background-position:60% center;}
    .vm-hero-copy-overlay{padding:2.35rem 1.45rem 1.85rem;max-width:94%;}
    .vm-kicker{margin-bottom:.8rem;}
    .vm-title{font-size:1.68rem;line-height:1.08;}
    .vm-sub{font-size:.88rem;line-height:1.48;}
    .vm-author{border-radius:12px;line-height:1.35;font-size:.74rem;}
    .vm-version{font-size:.68rem;}
}

.executive-panel{
    margin-top:.4rem;
    padding:1.15rem;
    border:1px solid #dbe5ef;
    border-radius:18px;
    background:linear-gradient(180deg,#fbfdff 0%,#f7fafc 100%);
    box-shadow:0 8px 24px rgba(15,23,42,.05);
}
.executive-grid{
    display:grid;
    grid-template-columns:repeat(4,minmax(0,1fr));
    gap:.8rem;
}
.executive-card{
    min-height:112px;
    padding:1rem 1.05rem;
    border:1px solid #dbe5ef;
    border-radius:14px;
    background:#fff;
    box-shadow:0 3px 10px rgba(15,23,42,.035);
}
.executive-card-label{
    color:#64748b;
    font-size:.74rem;
    font-weight:800;
    letter-spacing:.045em;
    text-transform:uppercase;
}
.executive-card-value{
    margin-top:.28rem;
    color:#0f172a;
    font-size:1.72rem;
    line-height:1.05;
    font-weight:850;
    letter-spacing:-.025em;
}
.executive-card-note{
    margin-top:.55rem;
    color:#64748b;
    font-size:.76rem;
    line-height:1.25;
    font-weight:650;
}
.executive-card-note.positive{color:#15803d;}
.executive-card-note.negative{color:#b91c1c;}
.executive-card-note.neutral{color:#64748b;}
.summary-card{
    margin-top:1rem;
    padding:1rem 1.1rem;
    border-left:4px solid #94a3b8;
    border-radius:0 12px 12px 0;
    background:#f8fafc;
    color:#334155;
    line-height:1.55;
    font-size:.93rem;
}
.summary-label{
    display:block;
    margin-bottom:.35rem;
    color:#0f172a;
    font-size:.76rem;
    font-weight:850;
    letter-spacing:.04em;
    text-transform:uppercase;
}
div[data-testid="stButton"] > button[kind="primary"]{
    min-height:3.25rem;
    border:1px solid rgba(79,70,229,.18);
    border-radius:13px;
    background:linear-gradient(135deg,#4f46e5 0%,#2563eb 48%,#14b8a6 100%);
    color:#fff;
    font-weight:850;
    font-size:1rem;
    letter-spacing:.01em;
    box-shadow:
        0 10px 24px rgba(37,99,235,.20),
        inset 0 1px 0 rgba(255,255,255,.18);
    transition:all .18s ease;
}
div[data-testid="stButton"] > button[kind="primary"]:hover{
    border-color:rgba(79,70,229,.28);
    background:linear-gradient(135deg,#4338ca 0%,#1d4ed8 48%,#0f9f91 100%);
    box-shadow:
        0 13px 28px rgba(37,99,235,.26),
        inset 0 1px 0 rgba(255,255,255,.20);
    transform:translateY(-1px);
}
div[data-testid="stButton"] > button[kind="primary"]:active{
    transform:translateY(0);
    box-shadow:0 7px 16px rgba(37,99,235,.18);
}
div[data-testid="stButton"] > button[kind="primary"]:disabled{
    border-color:#dbe3ec;
    background:linear-gradient(135deg,#edf2f7 0%,#e8eef5 100%);
    color:#9aa8b8;
    box-shadow:none;
    transform:none;
}
@media(max-width:900px){
    .executive-grid{grid-template-columns:repeat(2,minmax(0,1fr));}
}
@media(max-width:560px){
    .executive-grid{grid-template-columns:1fr;}
}
.quick-nav{
    display:flex;align-items:center;justify-content:space-between;gap:1rem;
    margin:.15rem 0 1rem;padding:.82rem 1rem;border:1px solid #dbe5ef;
    border-radius:14px;background:#fff;box-shadow:0 4px 16px rgba(15,23,42,.04);
    color:#334155;font-size:.92rem;
}
.quick-nav-badges{display:flex;gap:.45rem;flex-wrap:wrap;justify-content:flex-end;}
.quick-nav-badges span{padding:.28rem .58rem;border-radius:999px;background:#eff6ff;color:#1d4ed8;font-size:.73rem;font-weight:800;}
.section-kicker{margin-top:1.35rem;color:#2563eb;font-size:.72rem;font-weight:900;letter-spacing:.13em;}
.workflow-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.75rem;padding:.2rem 0 .45rem;}
.workflow-step{padding:1rem;border:1px solid #dbe5ef;border-radius:14px;background:#fbfdff;}
.workflow-step span{display:inline-flex;width:28px;height:28px;align-items:center;justify-content:center;border-radius:50%;background:#2563eb;color:#fff;font-weight:900;margin-bottom:.65rem;}
.workflow-step strong,.workflow-step small{display:block}.workflow-step small{margin-top:.35rem;color:#64748b;line-height:1.45;}
.guide-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:.7rem;margin:.35rem 0 .75rem;}
.guide-card{padding:.9rem;border:1px solid #dbe5ef;border-radius:13px;background:#fff;min-height:148px;}
.guide-card strong{color:#0f172a}.guide-card p{margin:.6rem 0 0;color:#475569;font-size:.82rem;line-height:1.4;}
.card-heading{display:flex;align-items:center;gap:.55rem;margin-bottom:.15rem}.mini-icon{display:inline-flex;width:30px;height:30px;align-items:center;justify-content:center;border-radius:9px;background:linear-gradient(135deg,#eff6ff,#ecfeff);border:1px solid #dbeafe;color:#2563eb;flex:0 0 auto}.mini-icon svg{width:17px;height:17px;stroke:currentColor;fill:none;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.guide-card code{font-size:.73rem;color:#1d4ed8;background:#eff6ff;padding:.13rem .28rem;border-radius:5px;}
.guide-note{padding:.72rem .85rem;border-radius:10px;background:#f8fafc;color:#64748b;font-size:.82rem;}
.demo-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.8rem;margin-top:.4rem;}
.demo-card{padding:.75rem;border:1px solid #dbe5ef;border-radius:14px;background:#fff;}
.demo-placeholder{display:flex;align-items:center;justify-content:center;min-height:128px;margin-bottom:.75rem;border:1px dashed #93c5fd;border-radius:10px;background:linear-gradient(135deg,#f8fbff,#eef6ff);color:#2563eb;text-align:center;font-size:.78rem;font-weight:800;}
.demo-card strong,.demo-card small{display:block}.demo-card small{margin-top:.3rem;color:#64748b;line-height:1.4}.demo-title{display:flex;align-items:center;gap:.5rem}.demo-title .mini-icon{width:28px;height:28px}.demo-title .mini-icon svg{width:16px;height:16px}
div[data-testid="stExpander"]{border:1px solid #dbe5ef!important;border-radius:13px!important;background:#fff!important;box-shadow:0 3px 12px rgba(15,23,42,.025);margin-bottom:.55rem;}
div[data-testid="stExpander"] summary{font-weight:800;color:#0f172a;}
div[data-testid="stFileUploaderDropzone"]{border:1.5px dashed #93c5fd;border-radius:16px;background:linear-gradient(180deg,#fbfdff,#f8fbff);padding:1rem;}
@media(max-width:1050px){.guide-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.demo-grid{grid-template-columns:repeat(2,minmax(0,1fr));}}
@media(max-width:760px){.quick-nav{align-items:flex-start;flex-direction:column}.workflow-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.quick-nav-badges{justify-content:flex-start;}}
@media(max-width:560px){.workflow-grid,.guide-grid,.demo-grid{grid-template-columns:1fr;}}
.workflow-icon{display:flex;width:44px;height:44px;align-items:center;justify-content:center;border-radius:12px;background:linear-gradient(135deg,#eff6ff,#ecfeff);color:#2563eb;margin-bottom:.7rem;border:1px solid #dbeafe}.workflow-icon svg{width:23px;height:23px;stroke:currentColor}.workflow-step{position:relative;overflow:hidden}.workflow-step:after{content:"";position:absolute;inset:auto 0 0 0;height:3px;background:linear-gradient(90deg,#2563eb,#14b8a6)}
.date-reminder{display:flex;gap:.6rem;align-items:center;margin:.2rem 0 .85rem;padding:.65rem .8rem;border:1px solid #bfdbfe;border-radius:10px;background:#f8fbff;color:#475569;font-size:.84rem;line-height:1.4}.date-reminder svg{flex:0 0 auto;width:18px;height:18px;stroke:#2563eb}.date-reminder strong{color:#1e3a8a}.date-reminder b{color:#2563eb}
.validation-shell{margin:.2rem 0 1rem;padding:1rem;border:1px solid #dbe5ef;border-radius:16px;background:linear-gradient(180deg,#fbfdff,#f8fafc);box-shadow:0 8px 22px rgba(15,23,42,.04)}.validation-summary{display:flex;align-items:center;justify-content:space-between;gap:1rem;margin-bottom:.8rem}.validation-title{font-weight:850;color:#0f172a}.validation-count{padding:.28rem .58rem;border-radius:999px;background:#e0f2fe;color:#0369a1;font-size:.75rem;font-weight:850}.validation-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.65rem}.validation-card{display:flex;gap:.7rem;align-items:flex-start;padding:.78rem .82rem;border:1px solid #dbe5ef;border-radius:12px;background:#fff}.validation-card.detected{border-color:#bbf7d0;background:#f0fdf4}.validation-card.missing{border-color:#fecaca;background:#fef2f2}.validation-card.duplicate{border-color:#fde68a;background:#fffbeb}.validation-badge{display:flex;width:28px;height:28px;align-items:center;justify-content:center;border-radius:9px;font-weight:900}.detected .validation-badge{background:#dcfce7;color:#15803d}.missing .validation-badge{background:#fee2e2;color:#b91c1c}.duplicate .validation-badge{background:#fef3c7;color:#a16207}.validation-meta strong,.validation-meta small{display:block}.validation-meta strong{font-size:.82rem;color:#0f172a}.validation-meta small{margin-top:.2rem;color:#64748b;font-size:.72rem;line-height:1.3;word-break:break-word}
div[data-testid="stFileUploader"]{padding:.35rem 0 .15rem}div[data-testid="stFileUploaderDropzone"]{min-height:170px;display:flex;align-items:center;justify-content:center;border:2px dashed #60a5fa!important;background:radial-gradient(circle at 50% 15%,rgba(59,130,246,.08),transparent 42%),linear-gradient(180deg,#fbfdff,#f8fbff)!important;box-shadow:inset 0 0 0 1px rgba(255,255,255,.7),0 8px 22px rgba(37,99,235,.06)}div[data-testid="stFileUploaderDropzone"] button{border:0!important;border-radius:10px!important;background:linear-gradient(135deg,#2563eb,#4f46e5)!important;color:#fff!important;font-weight:800!important;padding:.55rem 1rem!important;box-shadow:0 8px 18px rgba(37,99,235,.2)!important}div[data-testid="stFileUploaderDropzone"] button:hover{transform:translateY(-1px)}
@media(max-width:1000px){.validation-grid{grid-template-columns:repeat(2,minmax(0,1fr));}}@media(max-width:560px){.validation-grid{grid-template-columns:1fr;}}

.page-nav-wrap{margin:.1rem 0 .35rem;}
div[data-testid="stTabs"]{margin-top:.1rem;}
div[data-testid="stTabs"] [data-baseweb="tab-list"]{
    gap:.35rem;
    padding:.32rem;
    border:1px solid #dbe5ef;
    border-radius:14px;
    background:#f8fafc;
    box-shadow:0 4px 14px rgba(15,23,42,.04);
    overflow-x:auto;
}
div[data-testid="stTabs"] button[data-baseweb="tab"]{
    min-height:42px;
    padding:.55rem 1rem;
    border-radius:10px;
    color:#475569;
    font-size:.84rem;
    font-weight:750;
    white-space:nowrap;
}
div[data-testid="stTabs"] button[data-baseweb="tab"]:first-child{
    background:linear-gradient(135deg,#2563eb,#4f46e5);
    color:#fff;
    box-shadow:0 7px 16px rgba(37,99,235,.20);
}
div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"]{
    background:#fff;
    color:#1d4ed8;
    box-shadow:0 3px 10px rgba(15,23,42,.08);
}
div[data-testid="stTabs"] button[data-baseweb="tab"]:first-child[aria-selected="true"]{
    background:linear-gradient(135deg,#2563eb,#4f46e5);
    color:#fff;
}
div[data-testid="stTabs"] [data-baseweb="tab-highlight"]{display:none;}
.compact-start{margin:.25rem 0 .9rem;color:#64748b;font-size:.84rem;}
@media(max-width:620px){
    div[data-testid="stTabs"] [data-baseweb="tab-list"]{padding:.25rem;gap:.2rem;}
    div[data-testid="stTabs"] button[data-baseweb="tab"]{min-height:39px;padding:.48rem .72rem;font-size:.78rem;}
}

</style>
""", unsafe_allow_html=True)

render_platform_hero()

st.markdown('<div class="page-nav-wrap"></div>', unsafe_allow_html=True)
generate_tab, guide_tab, preview_tab = st.tabs(
    ["🚀 Generate Report", "Guide", "Report Preview"]
)

with guide_tab:
    st.subheader("How the workflow works")
    st.markdown(
        """
        <div class="workflow-grid">
            <div class="workflow-step"><div class="workflow-icon"><svg viewBox="0 0 24 24" fill="none" stroke-width="1.8"><path d="M7 3h7l4 4v14H7z"/><path d="M14 3v5h5"/><path d="M9.5 13h5M9.5 17h5"/></svg></div><strong>Export from Tableau</strong><small>Set every report to the same date range, then download the eight required cross tabs.</small></div>
            <div class="workflow-step"><div class="workflow-icon"><svg viewBox="0 0 24 24" fill="none" stroke-width="1.8"><path d="M12 16V4"/><path d="m7 9 5-5 5 5"/><path d="M5 20h14"/></svg></div><strong>Upload together</strong><small>Add all files in one step; the platform identifies and checks them automatically.</small></div>
            <div class="workflow-step"><div class="workflow-icon"><svg viewBox="0 0 24 24" fill="none" stroke-width="1.8"><path d="M12 3v3M12 18v3M3 12h3M18 12h3"/><circle cx="12" cy="12" r="4"/><path d="m5.6 5.6 2.1 2.1M16.3 16.3l2.1 2.1M18.4 5.6l-2.1 2.1M7.7 16.3l-2.1 2.1"/></svg></div><strong>Generate report</strong><small>Validate the source files and build the complete performance report.</small></div>
            <div class="workflow-step"><div class="workflow-icon"><svg viewBox="0 0 24 24" fill="none" stroke-width="1.8"><path d="M12 4v12"/><path d="m7 11 5 5 5-5"/><path d="M5 20h14"/></svg></div><strong>Download outputs</strong><small>Download the finished HTML, Excel and PDF reports directly.</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Required Tableau exports")
    st.caption("Open each Tableau report, select Download → Cross Tab, choose the listed view, and export in CSV format.")
    st.markdown(
        """
        <div class="date-reminder"><svg viewBox="0 0 24 24" fill="none" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="M12 10v5M12 7h.01"/></svg><div><strong>Date range reminder:</strong> Set the date control to <b>Range</b> and use the exact same start and end dates in every Tableau report.</div></div>
        <div class="guide-grid">
            <div class="guide-card">
                <div class="card-heading"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M5 3h14v18H5z"/><path d="M8 8h8M8 12h8M8 16h5"/></svg></span><strong>Executive Report</strong></div>
                <p><b>LFL</b> → <code>LFL.csv</code></p>
                <p><b>Group TY</b> → <code>Group by Site.csv</code></p>
            </div>
            <div class="guide-card">
                <div class="card-heading"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M4 19V9M10 19V5M16 19v-7M22 19V3"/></svg></span><strong>Group - Gender</strong></div>
                <p><b>Apparel</b> → <code>Apparel.csv</code></p>
                <p><b>Accessories</b> → <code>Accessories.csv</code></p>
                <p><b>Footwear</b> → <code>Footwear.csv</code></p>
            </div>
            <div class="guide-card">
                <div class="card-heading"><span class="mini-icon"><svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3"/><path d="M3 20c0-4 2.5-6 6-6s6 2 6 6"/><path d="M17 8h4M17 12h4M17 16h4"/></svg></span><strong>Group - Gender Details</strong></div>
                <p><b>Footwear → D Net Sales</b> → <code>D Net Sales.csv</code></p>
            </div>
            <div class="guide-card">
                <div class="card-heading"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M4 7h16M4 12h16M4 17h16"/><circle cx="8" cy="7" r="1.5"/><circle cx="15" cy="12" r="1.5"/><circle cx="11" cy="17" r="1.5"/></svg></span><strong>Option</strong></div>
                <p><b>Option List</b> → <code>Option List.csv</code></p>
            </div>
            <div class="guide-card">
                <div class="card-heading"><span class="mini-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg></span><strong>Traffic by Hour</strong></div>
                <p><b>Table</b> → <code>Table.csv</code></p>
            </div>
        </div>
        <div class="guide-note">The uploader checks file contents as well as filenames, so minor filename differences are usually accepted.</div>
        """,
        unsafe_allow_html=True,
    )

with preview_tab:
    st.subheader("Report preview")
    st.caption("An anonymized production example will be added after final cloud validation. These placeholders reserve the intended layout.")
    st.markdown(
        """
        <div class="demo-grid">
            <div class="demo-card"><div class="demo-placeholder">Production screenshot coming soon</div><div class="demo-title"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M4 20V10M10 20V4M16 20v-7M22 20V7"/></svg></span><strong>Executive Dashboard</strong></div><small>Headline KPIs, trends, sales mix and budget performance.</small></div>
            <div class="demo-card"><div class="demo-placeholder">Production screenshot coming soon</div><div class="demo-title"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z"/></svg></span><strong>Department Analysis</strong></div><small>Department and subdepartment performance with budget gaps.</small></div>
            <div class="demo-card"><div class="demo-placeholder">Production screenshot coming soon</div><div class="demo-title"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M20 13 11 4H4v7l9 9 7-7Z"/><circle cx="7.5" cy="7.5" r="1"/></svg></span><strong>Brands & Products</strong></div><small>Top brands, top sellers and contribution insights.</small></div>
            <div class="demo-card"><div class="demo-placeholder">Production screenshot coming soon</div><div class="demo-title"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M5 3h14v18H5z"/><path d="M8 8h8M8 12h8M8 16h5"/></svg></span><strong>Executive Summary</strong></div><small>Automated findings, opportunities and operational priorities.</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("#### Included in every report")
    st.markdown("""
- Executive KPI dashboard and automated summary
- Sales, budget attainment and budget-gap analysis
- Traffic, conversion, transactions, ATV and IPC / UPT
- Department and subdepartment performance
- Top brands and top-selling products
- Data validation checks and downloadable HTML, Excel and PDF outputs
""")


REQUIRED_UPLOADS: Dict[str, Dict[str, object]] = {
    key: {
        "label": definition["label"],
        "expected_file": INPUT_FILES[key],
    }
    for key, definition in REPORT_DEFINITIONS.items()
}


def safe_filename(text: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in ["-", "_"] else "_" for ch in text.strip())
    return cleaned.strip("_") or "VM_KPI_Report"


def get_metric(kpi_master: pd.DataFrame, name: str, default=0.0):
    row = kpi_master[kpi_master["kpi"].astype(str).str.lower() == name.lower()]
    if row.empty:
        return default
    return row.iloc[0].get("value", default)


def get_metric_field(
    kpi_master: pd.DataFrame,
    name: str,
    field: str,
    default=0.0,
):
    row = kpi_master[kpi_master["kpi"].astype(str).str.lower() == name.lower()]
    if row.empty:
        return default
    value = row.iloc[0].get(field, default)
    return default if pd.isna(value) else value


def embed_images_in_html(html_text: str) -> str:
    chart_dir = Path("assets/charts")
    if not chart_dir.exists():
        return html_text

    for image_path in chart_dir.glob("*.png"):
        encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
        data_uri = f"data:image/png;base64,{encoded}"
        html_text = html_text.replace(f"../assets/charts/{image_path.name}", data_uri)
        html_text = html_text.replace(f"assets/charts/{image_path.name}", data_uri)

    return html_text


def export_pdf(html_text: str) -> bytes:
    """Convert the completed HTML report into a downloadable PDF."""
    from weasyprint import HTML

    return HTML(string=html_text, base_url=str(APP_DIR)).write_pdf()


@st.cache_data(show_spinner=False)
def build_outputs_from_uploads(
    file_bytes: Dict[str, bytes],
    store_name: str,
    report_period: str,
) -> Tuple[bytes, bytes, bytes, bytes, Dict[str, pd.DataFrame], Dict[str, pd.DataFrame | str]]:
    with tempfile.TemporaryDirectory() as temp:
        base = Path(temp)

        for key, content in file_bytes.items():
            expected_name = str(REQUIRED_UPLOADS[key]["expected_file"])
            (base / expected_name).write_bytes(content)

        model_path = base / "vm_weekly_report_data_model.xlsx"
        calculations_path = base / "vm_weekly_report_calculations.xlsx"

        tables = build_model(base)

        if "report_info" in tables:
            tables["report_info"] = pd.DataFrame(
                [
                    ["Store Name", store_name],
                    ["Report Period", report_period],
                    ["Generated Date", date.today().isoformat()],
                    ["Data Model Version", "2.0 - combined department export"],
                ],
                columns=["field", "value"],
            )

        export_model(tables, model_path)

        outputs = build_report_calculations(model_path, store_name=store_name)
        export_calculations_to_excel(outputs, calculations_path)

        html_path = build_html_report(model_path, store_name=store_name, report_period=report_period)
        html_text = Path(html_path).read_text(encoding="utf-8")
        html_text = embed_images_in_html(html_text)
        html_bytes = html_text.encode("utf-8")

        pdf_bytes = export_pdf(html_text)

        return (
            model_path.read_bytes(),
            calculations_path.read_bytes(),
            html_bytes,
            pdf_bytes,
            tables,
            outputs,
        )


def show_validation(validation_df: pd.DataFrame) -> None:
    if validation_df.empty:
        st.warning("No validation checks were produced.")
        return

    passed_count = int(validation_df["passed"].sum()) if "passed" in validation_df.columns else 0
    total_count = len(validation_df)

    if passed_count == total_count:
        st.success(f"Validation passed: {passed_count}/{total_count} checks.")
    else:
        st.warning(
            f"Validation warning: {passed_count}/{total_count} checks passed. "
            "Review differences before sharing the report."
        )

    st.dataframe(validation_df, use_container_width=True, hide_index=True)


def show_kpi_cards(kpi_master: pd.DataFrame) -> None:
    if kpi_master.empty:
        st.warning("KPI table is empty.")
        return

    sales_vs_ly = float(get_metric_field(kpi_master, "Sales vs LY", "variance_pct", 0.0))
    budget_gap_pct = float(get_metric(kpi_master, "Budget Gap %", 0.0))
    transactions_vs_ly = float(
        get_metric_field(kpi_master, "Transactions", "variance_pct", 0.0)
    )
    ipc_vs_ly = float(
        get_metric_field(kpi_master, "IPC / UPT", "variance_pct", 0.0)
    )
    budget_attainment = float(get_metric(kpi_master, "Budget Attainment %", 0.0))

    def trend_note(value: float, comparison: str) -> tuple[str, str]:
        if value > 0:
            return f"▲ {pct(value)} {comparison}", "positive"
        if value < 0:
            return f"▼ {pct(abs(value))} {comparison}", "negative"
        return f"— {comparison}", "neutral"

    sales_note, sales_class = trend_note(sales_vs_ly, "vs LY")
    gap_note, gap_class = trend_note(budget_gap_pct, "vs budget")
    transaction_note, transaction_class = trend_note(transactions_vs_ly, "vs LY")
    ipc_note, ipc_class = trend_note(ipc_vs_ly, "vs LY")

    cards = [
        ("Net Sales", money(get_metric(kpi_master, "Net Sales")), sales_note, sales_class),
        ("Budget Gap", money(get_metric(kpi_master, "Budget Gap")), gap_note, gap_class),
        ("Conversion", pct(get_metric(kpi_master, "Conversion %")), "Transactions ÷ footfall", "neutral"),
        ("Footfall", number(get_metric(kpi_master, "Footfall")), "Store traffic", "neutral"),
        ("Transactions", number(get_metric(kpi_master, "Transactions")), transaction_note, transaction_class),
        ("ATV", money(get_metric(kpi_master, "ATV")), "Average transaction value", "neutral"),
        ("IPC / UPT", f"{get_metric(kpi_master, 'IPC / UPT'):.2f}", ipc_note, ipc_class),
        ("Budget Attainment", pct(budget_attainment), "Net sales ÷ budget", "positive" if budget_attainment >= 1 else "negative"),
    ]

    # Keep the HTML compact and unindented. Streamlit Markdown interprets
    # indented lines after blank lines as code blocks, which previously caused
    # most KPI cards to appear as visible HTML source.
    card_html = []
    for label, value, note, note_class in cards:
        safe_label = html.escape(str(label))
        safe_value = html.escape(str(value))
        safe_note = html.escape(str(note))
        card_html.append(
            f'<div class="executive-card">'
            f'<div class="executive-card-label">{safe_label}</div>'
            f'<div class="executive-card-value">{safe_value}</div>'
            f'<div class="executive-card-note {note_class}">{safe_note}</div>'
            f'</div>'
        )

    preview_html = (
        '<div class="executive-panel">'
        '<div class="executive-grid">'
        + "".join(card_html)
        + '</div>'
        '</div>'
    )
    st.markdown(preview_html, unsafe_allow_html=True)


def upload_fingerprint(uploaded_files, store_name: str, report_start: date, report_end: date, report_period: str) -> str:
    """Return a stable fingerprint so stale generated outputs are cleared when inputs change."""
    digest = hashlib.sha256()
    digest.update(store_name.strip().encode("utf-8"))
    digest.update(report_start.isoformat().encode("utf-8"))
    digest.update(report_end.isoformat().encode("utf-8"))
    digest.update(report_period.strip().encode("utf-8"))

    for uploaded_file in sorted(uploaded_files or [], key=lambda file: file.name.lower()):
        data = uploaded_file.getvalue()
        digest.update(uploaded_file.name.encode("utf-8"))
        digest.update(str(len(data)).encode("utf-8"))
        digest.update(hashlib.sha256(data).digest())

    return digest.hexdigest()


def clear_generated_report() -> None:
    for key in [
        "model_bytes",
        "calculations_bytes",
        "html_report_bytes",
        "pdf_report_bytes",
        "tables",
        "outputs",
    ]:
        st.session_state.pop(key, None)
    st.session_state["report_ready"] = False


def show_upload_validation(validation: Dict[str, object]) -> None:
    rows = validation_rows(validation)
    detected_count = sum(1 for row in rows if row["Status"] == "Detected")
    cards = []
    symbols = {"Detected": "✓", "Missing": "!", "Duplicate": "⚠"}
    classes = {"Detected": "detected", "Missing": "missing", "Duplicate": "duplicate"}

    for row in rows:
        status = row["Status"]
        uploaded = row.get("Uploaded file") or "Not uploaded"
        cards.append(
            f'<div class="validation-card {classes.get(status, "missing")}">'
            f'<div class="validation-badge">{symbols.get(status, "!")}</div>'
            f'<div class="validation-meta"><strong>{html.escape(str(row["Required report"]))}</strong>'
            f'<small>{html.escape(str(uploaded))}</small></div></div>'
        )

    st.markdown(
        '<div class="validation-shell">'
        '<div class="validation-summary">'
        '<div class="validation-title">Source-file validation</div>'
        f'<div class="validation-count">{detected_count} of {len(rows)} identified</div>'
        '</div><div class="validation-grid">' + ''.join(cards) + '</div></div>',
        unsafe_allow_html=True,
    )

    missing = validation.get("missing", [])
    duplicates = validation.get("duplicates", {})
    unrecognized = validation.get("unrecognized", [])
    errors = validation.get("errors", [])

    if missing:
        labels = [REPORT_DEFINITIONS[key]["label"] for key in missing]
        st.warning(f"Missing required files: {', '.join(labels)}")

    if duplicates:
        for key, names in duplicates.items():
            st.warning(
                f"Duplicate detected for {REPORT_DEFINITIONS[key]['label']}: {', '.join(names)}. "
                "Remove one of these files before generating the report."
            )

    if errors:
        for error in errors:
            st.error(f"Could not read {error['filename']}: {error['error']}")

    if unrecognized:
        st.info("Additional unrecognized files will be ignored: " + ", ".join(unrecognized))

    if validation.get("ready"):
        st.success("All eight Tableau exports were identified successfully.")



with generate_tab:
    st.markdown(
        '<div class="compact-start"><strong>Start here:</strong> add the report details, upload the eight Tableau exports and generate the report.</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="section-kicker">BUILD YOUR REPORT</div>', unsafe_allow_html=True)
    st.subheader("Report setup")
    st.caption("Add the location and reporting period that should appear in the report header.")

    with st.container(border=True):
        setup_col1, setup_col2, setup_col3 = st.columns([1.35, 1, 1])
        with setup_col1:
            store_name = st.text_input(
                "Store / location",
                value="",
                placeholder="e.g., Sample Store",
                help="Required. The Tableau exports do not reliably contain the store name.",
            )
        with setup_col2:
            report_start = st.date_input("Report start date", value=date.today())
        with setup_col3:
            report_end = st.date_input("Report end date", value=date.today())

        default_report_period = f"{report_start.strftime('%b %d, %Y')} – {report_end.strftime('%b %d, %Y')}"
        report_period = st.text_input(
            "Report label / period",
            value=default_report_period,
            help="This text appears on the report header and can be edited manually.",
        )

    st.subheader("1. Upload Tableau exports")
    st.write("Drag and drop all eight required exports below. The platform will identify and validate each report automatically.")

    bulk_uploads = st.file_uploader(
        "Upload all Tableau exports",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="bulk_tableau_uploads",
        help="You may select all eight files at once. Original filenames can vary because the app also checks their contents.",
    )

    validation = classify_uploaded_files(bulk_uploads or [])

    st.subheader("2. Report validation")
    show_upload_validation(validation)

    current_fingerprint = upload_fingerprint(
        bulk_uploads or [], store_name, report_start, report_end, report_period
    )
    if st.session_state.get("input_fingerprint") != current_fingerprint:
        clear_generated_report()
        st.session_state["input_fingerprint"] = current_fingerprint

    location_missing = not store_name.strip()
    dates_invalid = report_start > report_end
    period_missing = not report_period.strip()

    if location_missing:
        st.warning("Enter the store / location before generating the report.")
    if dates_invalid:
        st.warning("The report start date cannot be after the report end date.")
    if period_missing:
        st.warning("Enter a report label / period before generating the report.")

    ready_to_generate = bool(
        validation.get("ready")
        and not location_missing
        and not dates_invalid
        and not period_missing
    )

    st.subheader("3. Generate report")
    if not ready_to_generate:
        st.warning("Missing required files or report details. Complete the items above to continue.")

    build_clicked = st.button(
        "Generate Report",
        type="primary",
        use_container_width=True,
        disabled=not ready_to_generate,
    )

    if build_clicked:
        with st.spinner("Building data model, calculations, charts and report..."):
            try:
                (
                    model_bytes,
                    calculations_bytes,
                    html_report_bytes,
                    pdf_report_bytes,
                    tables,
                    outputs,
                ) = build_outputs_from_uploads(
                    validation["file_bytes"],
                    store_name.strip(),
                    report_period.strip(),
                )

                st.session_state["model_bytes"] = model_bytes
                st.session_state["calculations_bytes"] = calculations_bytes
                st.session_state["html_report_bytes"] = html_report_bytes
                st.session_state["pdf_report_bytes"] = pdf_report_bytes
                st.session_state["tables"] = tables
                st.session_state["outputs"] = outputs
                st.session_state["report_ready"] = True
                st.success("Report generated successfully.")

            except Exception as exc:
                clear_generated_report()
                st.error("The report could not be generated.")
                st.exception(exc)

    if st.session_state.get("report_ready"):
        outputs = st.session_state["outputs"]

        st.subheader("4. Executive preview")
        show_kpi_cards(outputs["kpi_master"])

        summary_text = outputs.get("executive_summary_text", "")
        if summary_text:
            safe_summary = html.escape(str(summary_text)).replace("\n", "<br>")
            st.markdown(
                f"""
                <div class="summary-card">
                    <span class="summary-label">Auto executive summary</span>
                    {safe_summary}
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.subheader("5. Data validation checks")
        show_validation(outputs["validation"])

        st.subheader("6. Report tables")

        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            ["Department KPIs", "Subdepartment Gaps", "Opportunities", "Top Brands", "Traffic"]
        )

        with tab1:
            st.dataframe(outputs["department_kpis"], use_container_width=True, hide_index=True)

        with tab2:
            st.dataframe(outputs["sub_department_kpis"], use_container_width=True, hide_index=True)

        with tab3:
            st.caption(
                "Priority Score = absolute budget gap % weighted by sales mix. "
                "It ranks areas that are both underperforming and meaningful to the business."
            )
            st.dataframe(outputs["opportunity_ranking"], use_container_width=True, hide_index=True)

        with tab4:
            st.dataframe(outputs["top_brands_fixed"], use_container_width=True, hide_index=True)

        with tab5:
            st.dataframe(outputs["traffic_summary"], use_container_width=True, hide_index=True)
            st.dataframe(outputs["hourly_traffic"], use_container_width=True, hide_index=True)

        st.subheader("7. Download outputs")

        # The requested convention uses the report end date as the report date.
        report_date_text = report_end.strftime("%Y-%m-%d")
        base_name = safe_filename(f"{store_name}_VM_KPI_Report_{report_date_text}")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.download_button(
                "Download data model",
                data=st.session_state["model_bytes"],
                file_name=f"{base_name}_Data_Model.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col2:
            st.download_button(
                "Download calculations",
                data=st.session_state["calculations_bytes"],
                file_name=f"{base_name}_Calculations.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col3:
            st.download_button(
                "Download HTML report",
                data=st.session_state["html_report_bytes"],
                file_name=f"{base_name}.html",
                mime="text/html",
                use_container_width=True,
            )

        with col4:
            st.download_button(
                "Download PDF report",
                data=st.session_state["pdf_report_bytes"],
                file_name=f"{base_name}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
