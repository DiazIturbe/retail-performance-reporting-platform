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
from PIL import Image

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
EXAMPLE_ASSET_DIR = ASSETS_DIR / "example_report"
SAMPLE_REPORT_PDF = EXAMPLE_ASSET_DIR / "Synthetic_Store_Performance_Report.pdf"
SAMPLE_REPORT_HTML = EXAMPLE_ASSET_DIR / "Synthetic_Store_Performance_Report.html"
REPORT_TEMPLATE_VERSION = "2026-07-30.1"

try:
    PAGE_ICON = Image.open(FAVICON_PATH) if FAVICON_PATH.exists() else "📈"
except Exception:
    PAGE_ICON = "📈"

st.set_page_config(
    page_title="Store Performance Reporting Platform",
    page_icon=PAGE_ICON,
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
                        Created by Diego Diaz Iturbe · Retail Analytics · Reporting Automation
                    </div>
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
    min-height:285px;
    margin:0 0 1.15rem 0;
    border-radius:0 0 24px 24px;
    color:white;
    box-shadow:0 10px 28px rgba(15,23,42,.12);
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
    padding:2.35rem 3rem 2.15rem;
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
    .vm-hero{min-height:265px;border-radius:0 0 18px 18px;}
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

 .executive-panel{margin-top:.3rem;padding:.7rem;border:1px solid #dbe5ef;border-radius:15px;background:linear-gradient(180deg,#fbfdff 0%,#f7fafc 100%);box-shadow:0 5px 16px rgba(15,23,42,.04)}
.executive-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.5rem}.executive-card{min-height:84px;padding:.68rem .72rem;border:1px solid #dbe5ef;border-radius:11px;background:#fff;box-shadow:0 2px 7px rgba(15,23,42,.025)}.executive-card-label{color:#64748b;font-size:.65rem;font-weight:800;letter-spacing:.04em;text-transform:uppercase}.executive-card-value{margin-top:.16rem;color:#0f172a;font-size:1.38rem;line-height:1.03;font-weight:850;letter-spacing:-.022em}.executive-card-note{margin-top:.32rem;color:#64748b;font-size:.65rem;line-height:1.2;font-weight:650}.executive-card-note.positive{color:#15803d}.executive-card-note.negative{color:#b91c1c}.executive-card-note.neutral{color:#64748b}.summary-card{margin-top:.65rem;padding:.72rem .82rem;border-left:3px solid #94a3b8;border-radius:0 10px 10px 0;background:#f8fafc;color:#334155;line-height:1.42;font-size:.82rem}.summary-label{display:block;margin-bottom:.22rem;color:#0f172a;font-size:.67rem;font-weight:850;letter-spacing:.04em;text-transform:uppercase}.report-ready-panel{margin:.35rem 0 .65rem;padding:.78rem;border:1px solid #bfdbfe;border-radius:14px;background:linear-gradient(135deg,#eff6ff,#f8fbff 55%,#f0fdfa)}.report-ready-title{font-size:1rem;font-weight:850;color:#0f172a}.report-ready-copy{margin-top:.18rem;color:#526174;font-size:.78rem;line-height:1.4}.format-note{margin:.2rem 0 .45rem;color:#64748b;font-size:.72rem}.recommended-tag{display:inline-block;margin-left:.3rem;padding:.12rem .34rem;border-radius:999px;background:#dbeafe;color:#1d4ed8;font-size:.56rem;font-weight:850;text-transform:uppercase}.st-key-html_report_download button{min-height:2.8rem!important;border-radius:10px!important;border-color:#2563eb!important;background:linear-gradient(135deg,#2563eb,#1d4ed8)!important;color:#fff!important;font-weight:800!important;box-shadow:0 7px 16px rgba(37,99,235,.16)!important}.st-key-pdf_report_download button{min-height:2.8rem!important;border-radius:10px!important;border:1px solid #94a3b8!important;background:#fff!important;color:#263244!important;font-weight:780!important}@media(max-width:900px){.executive-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:560px){.executive-panel{padding:.5rem}.executive-grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:.4rem}.executive-card{min-height:76px;padding:.55rem .58rem}.executive-card-label{font-size:.58rem}.executive-card-value{font-size:1.13rem}.executive-card-note{font-size:.57rem;margin-top:.24rem}.summary-card{padding:.6rem .68rem;font-size:.74rem;line-height:1.38}}
.quick-nav{
    display:flex;align-items:center;justify-content:space-between;gap:1rem;
    margin:.15rem 0 1rem;padding:.82rem 1rem;border:1px solid #dbe5ef;
    border-radius:14px;background:#fff;box-shadow:0 4px 16px rgba(15,23,42,.04);
    color:#334155;font-size:.92rem;
}
.quick-nav-badges{display:flex;gap:.45rem;flex-wrap:wrap;justify-content:flex-end;}
.quick-nav-badges span{padding:.28rem .58rem;border-radius:999px;background:#eff6ff;color:#1d4ed8;font-size:.73rem;font-weight:800;}
.section-kicker{margin-top:1.35rem;color:#2563eb;font-size:.72rem;font-weight:900;letter-spacing:.13em;}
.process-strip{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.55rem;margin:.55rem 0 1.25rem;}
.process-item{position:relative;padding:.72rem .8rem;border:1px solid #dbe5ef;border-radius:11px;background:#fff;color:#475569;font-size:.78rem;font-weight:750;line-height:1.35;}
.process-item b{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;margin-right:.42rem;border-radius:50%;background:#eaf3ff;color:#1d4ed8;font-size:.72rem;}
.process-item:not(:last-child):after{content:"";position:absolute;right:-.42rem;top:50%;width:.3rem;border-top:1px solid #93c5fd;}
.upload-guidance{margin:.25rem 0 .8rem;padding:.72rem .85rem;border-left:3px solid #2563eb;border-radius:0 10px 10px 0;background:#f8fbff;color:#475569;font-size:.84rem;line-height:1.5;}
.upload-guidance strong{color:#1e3a8a;}
.about-panel{margin-top:1.25rem;padding:1rem 1.05rem;border:1px solid #dbe5ef;border-radius:14px;background:#f8fafc;color:#475569;font-size:.88rem;line-height:1.55;}
.about-panel strong{display:block;margin-bottom:.25rem;color:#0f172a;font-size:.96rem;}
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
.demo-image{display:block;width:100%;aspect-ratio:16/9;object-fit:cover;object-position:top left;margin-bottom:.75rem;border:1px solid #dbe5ef;border-radius:10px;background:#f8fafc;box-shadow:0 5px 14px rgba(15,23,42,.07);}
.demo-card strong,.demo-card small{display:block}.demo-card small{margin-top:.3rem;color:#64748b;line-height:1.4}.demo-title{display:flex;align-items:center;gap:.5rem}.demo-title .mini-icon{width:28px;height:28px}.demo-title .mini-icon svg{width:16px;height:16px}
div[data-testid="stExpander"]{border:1px solid #dbe5ef!important;border-radius:13px!important;background:#fff!important;box-shadow:0 3px 12px rgba(15,23,42,.025);margin-bottom:.55rem;}
div[data-testid="stExpander"] summary{font-weight:800;color:#0f172a;}
div[data-testid="stFileUploaderDropzone"]{border:1.5px dashed #93c5fd;border-radius:16px;background:linear-gradient(180deg,#fbfdff,#f8fbff);padding:1rem;}
@media(max-width:1050px){.guide-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.demo-grid{grid-template-columns:repeat(2,minmax(0,1fr));}}
@media(max-width:760px){.quick-nav{align-items:flex-start;flex-direction:column}.workflow-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.quick-nav-badges{justify-content:flex-start;}}
@media(max-width:560px){.workflow-grid,.guide-grid,.demo-grid{grid-template-columns:1fr;}}
@media(max-width:760px){.process-strip{grid-template-columns:repeat(2,minmax(0,1fr));}.process-item:after{display:none;}}
@media(max-width:480px){.process-strip{grid-template-columns:1fr;}}
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


/* ------------------------------------------------------------------
   Responsive density improvements
   Preserve the desktop workflow while making the mobile experience
   faster to scan and less vertically demanding.
   ------------------------------------------------------------------ */

/* Slightly tighter rhythm on all screen sizes. */
.block-container{
    padding-top:.55rem;
    padding-bottom:2.25rem;
}

h1, h2, h3{
    letter-spacing:-.02em;
}

h2, div[data-testid="stHeadingWithActionElements"] h2{
    margin-top:1.2rem;
    margin-bottom:.45rem;
}

p, .stCaption{
    line-height:1.45;
}

/* Streamlit alerts: noticeable, but no longer visually dominant. */
div[data-testid="stAlert"]{
    padding:.68rem .82rem;
    margin:.35rem 0 .55rem;
    border-radius:11px;
}
div[data-testid="stAlert"] [data-testid="stMarkdownContainer"] p{
    margin:0;
    font-size:.88rem;
    line-height:1.38;
}
div[data-testid="stAlert"] [data-testid="stAlertContentInfo"],
div[data-testid="stAlert"] [data-testid="stAlertContentWarning"],
div[data-testid="stAlert"] [data-testid="stAlertContentError"],
div[data-testid="stAlert"] [data-testid="stAlertContentSuccess"]{
    gap:.55rem;
}

/* Make bordered form containers more efficient without shrinking controls. */
div[data-testid="stVerticalBlockBorderWrapper"]{
    border-radius:14px;
}
div[data-testid="stVerticalBlockBorderWrapper"] > div{
    padding:.82rem .9rem;
}
div[data-testid="stTextInput"] input,
div[data-testid="stDateInput"] input{
    min-height:44px;
}

/* Compact the explanatory and validation elements on desktop too. */
.process-strip{
    gap:.45rem;
    margin:.35rem 0 .9rem;
}
.process-item{
    padding:.58rem .68rem;
    font-size:.75rem;
}
.process-item b{
    width:20px;
    height:20px;
    margin-right:.34rem;
}
.upload-guidance{
    margin:.15rem 0 .6rem;
    padding:.62rem .75rem;
    font-size:.80rem;
    line-height:1.42;
}
.validation-shell{
    margin:.1rem 0 .7rem;
    padding:.78rem;
    border-radius:14px;
}
.validation-summary{
    margin-bottom:.58rem;
}
.validation-grid{
    gap:.48rem;
}
.validation-card{
    gap:.55rem;
    padding:.62rem .68rem;
    border-radius:10px;
}
.validation-badge{
    width:25px;
    height:25px;
    border-radius:8px;
}
.validation-meta strong{
    font-size:.79rem;
    line-height:1.25;
}
.validation-meta small{
    margin-top:.12rem;
    font-size:.69rem;
}
div[data-testid="stFileUploader"]{
    padding:.15rem 0 .05rem;
}
div[data-testid="stFileUploaderDropzone"]{
    min-height:128px;
}

@media(max-width:620px){
    /* Keep the product identity, but replace the poster-like hero with a
       compact mobile header. */
    .block-container{
        padding-top:.25rem;
        padding-left:.85rem;
        padding-right:.85rem;
        padding-bottom:1.4rem;
    }
    .vm-hero{
        min-height:132px;
        margin:0 0 .55rem;
        border-radius:0 0 15px 15px;
        box-shadow:0 5px 16px rgba(15,23,42,.10);
    }
    .vm-hero-image{
        background-position:52% center;
    }
    .vm-hero-overlay{
        background:linear-gradient(90deg,rgba(3,12,30,.93),rgba(4,19,45,.70) 68%,rgba(4,18,40,.28));
    }
    .vm-hero-copy-overlay,
    .vm-hero-fallback{
        padding:1rem 1.05rem;
        max-width:100%;
    }
    .vm-kicker{
        margin-bottom:.3rem;
        font-size:.60rem;
        letter-spacing:.13em;
    }
    .vm-title{
        max-width:92%;
        margin-bottom:0;
        font-size:1.28rem;
        line-height:1.08;
    }
    .vm-sub,
    .vm-author{
        display:none;
    }

    /* Tabs remain easy to tap while using less space. */
    div[data-testid="stTabs"]{
        margin-top:0;
    }
    div[data-testid="stTabs"] [data-baseweb="tab-list"]{
        margin-bottom:.35rem;
        border-radius:11px;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"]{
        min-height:38px;
        padding:.42rem .56rem;
        font-size:.74rem;
    }

    /* Mobile headings: strong hierarchy without poster-sized type. */
    h1{font-size:1.65rem!important;line-height:1.12!important;}
    h2, div[data-testid="stHeadingWithActionElements"] h2{
        margin-top:.85rem!important;
        margin-bottom:.32rem!important;
        font-size:1.42rem!important;
        line-height:1.16!important;
    }
    h3{font-size:1.08rem!important;line-height:1.2!important;}
    .stCaption, div[data-testid="stCaptionContainer"]{
        font-size:.78rem!important;
        line-height:1.38!important;
    }

    /* One compact workflow row rather than four large cards. */
    .process-strip{
        grid-template-columns:repeat(4,minmax(0,1fr));
        gap:.28rem;
        margin:.15rem 0 .65rem;
    }
    .process-item{
        display:flex;
        flex-direction:column;
        align-items:center;
        justify-content:center;
        min-height:54px;
        padding:.36rem .18rem;
        text-align:center;
        font-size:.60rem;
        line-height:1.18;
        border-radius:9px;
    }
    .process-item b{
        width:18px;
        height:18px;
        margin:0 0 .2rem;
        font-size:.62rem;
    }
    .process-item:after{display:none!important;}

    /* Form card and widgets stay touch-friendly but lose excess whitespace. */
    div[data-testid="stVerticalBlockBorderWrapper"] > div{
        padding:.62rem .68rem;
    }
    div[data-testid="stVerticalBlock"]{
        gap:.55rem;
    }
    div[data-testid="stTextInput"],
    div[data-testid="stDateInput"]{
        margin-bottom:.1rem;
    }
    div[data-testid="stTextInput"] input,
    div[data-testid="stDateInput"] input{
        min-height:44px;
        padding:.55rem .72rem;
        font-size:.88rem;
    }
    label[data-testid="stWidgetLabel"] p{
        font-size:.82rem;
        line-height:1.25;
    }

    /* Shorter guidance panel. The secondary data-handling note remains
       available on desktop, while mobile focuses on the immediate task. */
    .upload-guidance{
        margin:.08rem 0 .45rem;
        padding:.55rem .62rem;
        font-size:.76rem;
        line-height:1.36;
        border-radius:0 8px 8px 0;
    }
    .upload-guidance small{
        display:none;
    }

    div[data-testid="stFileUploaderDropzone"]{
        min-height:88px;
        padding:.5rem!important;
        border-width:1px!important;
    }
    div[data-testid="stFileUploaderDropzone"] button{
        padding:.48rem .72rem!important;
        font-size:.78rem!important;
    }

    /* Validation becomes a dense status list rather than a stack of cards. */
    .validation-shell{
        margin:.08rem 0 .55rem;
        padding:.58rem;
        border-radius:12px;
    }
    .validation-summary{
        gap:.5rem;
        margin-bottom:.45rem;
    }
    .validation-title{font-size:.86rem;}
    .validation-count{
        padding:.22rem .46rem;
        font-size:.68rem;
    }
    .validation-grid{
        grid-template-columns:1fr;
        gap:.34rem;
    }
    .validation-card{
        align-items:center;
        gap:.48rem;
        min-height:50px;
        padding:.43rem .5rem;
        border-radius:9px;
    }
    .validation-badge{
        width:23px;
        height:23px;
        flex:0 0 23px;
        font-size:.76rem;
    }
    .validation-meta strong{
        font-size:.75rem;
        line-height:1.2;
    }
    .validation-meta small{
        margin-top:.06rem;
        font-size:.64rem;
        line-height:1.18;
    }

    /* Compact routine system feedback while retaining readable contrast and
       comfortable line length. */
    div[data-testid="stAlert"]{
        padding:.52rem .62rem;
        margin:.22rem 0 .38rem;
        border-radius:9px;
    }
    div[data-testid="stAlert"] [data-testid="stMarkdownContainer"] p{
        font-size:.76rem;
        line-height:1.32;
    }

    div[data-testid="stButton"] > button[kind="primary"]{
        min-height:46px;
        font-size:.88rem;
        border-radius:11px;
    }
}



/* Final density pass: compact report details and source validation. */
div[data-testid="stVerticalBlockBorderWrapper"] > div{
    padding:.68rem .78rem;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"]{
    gap:.34rem;
}
div[data-testid="stVerticalBlockBorderWrapper"] label[data-testid="stWidgetLabel"]{
    margin-bottom:.12rem;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stTextInput"],
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stDateInput"]{
    margin-bottom:0;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stCaptionContainer"]{
    margin-top:-.08rem;
    margin-bottom:.05rem;
    color:#64748b;
    font-size:.72rem;
}
/* The optional display label is subtly distinguished because it controls
   the exact wording used in the report header. */
input[aria-label="Report label (optional — shown in report)"]{
    background:#f8fbff!important;
    border-color:#bfdbfe!important;
    box-shadow:inset 0 0 0 1px rgba(59,130,246,.05)!important;
}
label:has(+ div input[aria-label="Report label (optional — shown in report)"]) p{
    color:#334155!important;
}
.validation-shell{
    padding:.58rem;
    margin:.05rem 0 .55rem;
}
.validation-summary{
    margin-bottom:.4rem;
}
.validation-grid{
    gap:.34rem;
}
.validation-card{
    min-height:0;
    padding:.46rem .52rem;
    gap:.44rem;
}
.validation-badge{
    width:22px;
    height:22px;
    flex:0 0 22px;
    border-radius:7px;
    font-size:.72rem;
}
.validation-meta strong{
    font-size:.75rem;
    line-height:1.18;
}
.validation-meta small{
    margin-top:.04rem;
    font-size:.64rem;
    line-height:1.14;
}

@media(max-width:620px){
    div[data-testid="stVerticalBlockBorderWrapper"] > div{
        padding:.48rem .54rem;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"]{
        gap:.26rem;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stHorizontalBlock"]{
        gap:.42rem;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] label[data-testid="stWidgetLabel"] p{
        font-size:.76rem;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stCaptionContainer"]{
        font-size:.66rem!important;
        line-height:1.25!important;
    }
    .validation-shell{
        padding:.42rem;
        margin:.04rem 0 .42rem;
    }
    .validation-summary{
        margin-bottom:.3rem;
    }
    .validation-grid{
        gap:.24rem;
    }
    .validation-card{
        min-height:42px;
        padding:.32rem .38rem;
        gap:.38rem;
        border-radius:8px;
    }
    .validation-badge{
        width:20px;
        height:20px;
        flex-basis:20px;
        font-size:.66rem;
    }
    .validation-meta strong{
        font-size:.71rem;
    }
    .validation-meta small{
        font-size:.60rem;
    }
}


/* Final mobile-form correction: denser inputs and dates kept side by side. */
div[data-testid="stVerticalBlockBorderWrapper"] > div{
    padding:.44rem .56rem!important;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"]{
    gap:.16rem!important;
}
div[data-testid="stVerticalBlockBorderWrapper"] label[data-testid="stWidgetLabel"]{
    margin-bottom:.02rem!important;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stTextInput"],
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stDateInput"]{
    margin-top:0!important;
    margin-bottom:0!important;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-baseweb="input"]{
    min-height:40px!important;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stTextInput"] input,
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stDateInput"] input{
    min-height:40px!important;
    padding:.38rem .62rem!important;
    line-height:1.2!important;
}

/* Streamlit normally stacks columns on narrow screens. The two report-date
   controls are short enough to remain usable as a two-column mobile row. */
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stHorizontalBlock"]{
    display:flex!important;
    flex-direction:row!important;
    flex-wrap:nowrap!important;
    align-items:flex-start!important;
    gap:.42rem!important;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]{
    flex:1 1 0!important;
    width:0!important;
    min-width:0!important;
}

/* Source-file validation is a status list, not a content panel: keep it terse. */
.validation-shell{
    padding:.34rem!important;
    margin:.02rem 0 .38rem!important;
    border-radius:11px!important;
}
.validation-summary{
    margin-bottom:.22rem!important;
    gap:.4rem!important;
}
.validation-grid{
    gap:.20rem!important;
}
.validation-card{
    min-height:0!important;
    padding:.28rem .34rem!important;
    gap:.32rem!important;
    border-radius:7px!important;
}
.validation-badge{
    width:18px!important;
    height:18px!important;
    flex:0 0 18px!important;
    border-radius:6px!important;
    font-size:.61rem!important;
}
.validation-meta strong{
    font-size:.69rem!important;
    line-height:1.12!important;
}
.validation-meta small{
    margin-top:0!important;
    font-size:.57rem!important;
    line-height:1.08!important;
}

@media(max-width:620px){
    div[data-testid="stVerticalBlockBorderWrapper"] > div{
        padding:.34rem .40rem!important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"]{
        gap:.11rem!important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stHorizontalBlock"]{
        gap:.34rem!important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] label[data-testid="stWidgetLabel"] p{
        font-size:.72rem!important;
        line-height:1.16!important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-baseweb="input"],
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stTextInput"] input,
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stDateInput"] input{
        min-height:39px!important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stTextInput"] input,
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stDateInput"] input{
        padding:.32rem .48rem!important;
        font-size:.78rem!important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stCaptionContainer"]{
        margin-top:-.02rem!important;
        font-size:.62rem!important;
        line-height:1.18!important;
    }
    .validation-shell{
        padding:.26rem!important;
        margin:.02rem 0 .30rem!important;
    }
    .validation-summary{
        margin-bottom:.17rem!important;
    }
    .validation-grid{
        gap:.16rem!important;
    }
    .validation-card{
        min-height:34px!important;
        padding:.22rem .28rem!important;
        gap:.27rem!important;
    }
    .validation-badge{
        width:17px!important;
        height:17px!important;
        flex-basis:17px!important;
        font-size:.57rem!important;
    }
    .validation-meta strong{font-size:.66rem!important;}
    .validation-meta small{font-size:.54rem!important;}
}



/* Report-details layout: authoritative compact rules. */
.st-key-report_details_card{
    padding:.48rem .62rem!important;
}
.st-key-report_details_card div[data-testid="stVerticalBlock"]{
    gap:.28rem!important;
}
.st-key-report_details_card label[data-testid="stWidgetLabel"]{
    margin-bottom:0!important;
}
.st-key-report_details_card label[data-testid="stWidgetLabel"] p{
    margin-bottom:0!important;
    font-size:.86rem!important;
    line-height:1.2!important;
}
.st-key-report_details_card div[data-baseweb="input"]{
    min-height:42px!important;
}
.st-key-report_details_card input{
    min-height:42px!important;
    padding:.34rem .58rem!important;
}

/* Keep the two date controls in a true two-column row on every viewport. */
.st-key-report_date_row div[data-testid="stHorizontalBlock"]{
    display:grid!important;
    grid-template-columns:minmax(0,1fr) minmax(0,1fr)!important;
    gap:.48rem!important;
    width:100%!important;
    align-items:start!important;
}
.st-key-report_date_row div[data-testid="stElementContainer"]{
    width:100%!important;
    min-width:0!important;
}
.st-key-report_date_row div[data-testid="stDateInput"]{
    width:100%!important;
    min-width:0!important;
}
.st-key-report_label_area{
    margin-top:.08rem!important;
    padding:.28rem .36rem .16rem!important;
    border:1px solid #e3ebf5!important;
    border-radius:9px!important;
    background:#f8fbff!important;
}
.st-key-report_label_area div[data-testid="stCaptionContainer"]{
    margin-top:-.08rem!important;
    color:#718096!important;
    font-size:.72rem!important;
    line-height:1.25!important;
}

@media(max-width:620px){
    .st-key-report_details_card{
        padding:.34rem .42rem!important;
    }
    .st-key-report_details_card div[data-testid="stVerticalBlock"]{
        gap:.16rem!important;
    }
    .st-key-report_details_card label[data-testid="stWidgetLabel"] p{
        font-size:.72rem!important;
        line-height:1.12!important;
    }
    .st-key-report_details_card div[data-baseweb="input"],
    .st-key-report_details_card input{
        min-height:38px!important;
    }
    .st-key-report_details_card input{
        padding:.26rem .42rem!important;
        font-size:.76rem!important;
    }
    .st-key-report_date_row div[data-testid="stHorizontalBlock"]{
        grid-template-columns:minmax(0,1fr) minmax(0,1fr)!important;
        gap:.34rem!important;
    }
    .st-key-report_label_area{
        padding:.20rem .28rem .10rem!important;
        border-radius:8px!important;
    }
    .st-key-report_label_area div[data-testid="stCaptionContainer"]{
        font-size:.62rem!important;
        line-height:1.16!important;
    }
}


/* Completion-flow polish: one clear success state and stronger download hierarchy. */
.report-complete-card{
    display:flex;
    align-items:flex-start;
    gap:.72rem;
    margin:.4rem 0 .75rem;
    padding:.82rem .9rem;
    border:1px solid #bbf7d0;
    border-radius:14px;
    background:linear-gradient(135deg,#f0fdf4 0%,#f8fffb 56%,#eff6ff 100%);
    box-shadow:0 5px 16px rgba(15,23,42,.045);
}
.report-complete-icon{
    display:flex;
    align-items:center;
    justify-content:center;
    width:2rem;
    height:2rem;
    flex:0 0 2rem;
    border-radius:999px;
    background:#dcfce7;
    color:#15803d;
    font-size:1rem;
    font-weight:900;
}
.report-complete-title{color:#0f172a;font-size:1rem;font-weight:850;line-height:1.2}
.report-complete-copy{margin-top:.16rem;color:#526174;font-size:.76rem;line-height:1.38}
.download-hero{
    margin:.15rem 0 .65rem;
    padding:.9rem;
    border:1px solid #bfdbfe;
    border-radius:15px;
    background:linear-gradient(145deg,#f8fbff 0%,#eff6ff 100%);
    box-shadow:0 8px 20px rgba(37,99,235,.07);
}
.download-hero-top{display:flex;align-items:center;justify-content:space-between;gap:.6rem;margin-bottom:.18rem}
.download-format-title{display:flex;align-items:center;gap:.42rem;color:#0f172a;font-size:.95rem;font-weight:850}
.download-format-icon{font-size:1.05rem;line-height:1}
.download-feature-list{display:flex;flex-wrap:wrap;gap:.28rem .6rem;margin:.45rem 0 .58rem;color:#526174;font-size:.68rem}
.download-feature{white-space:nowrap}
.download-feature::before{content:'✓';margin-right:.2rem;color:#15803d;font-weight:900}
.pdf-download-row{
    margin:.2rem 0 .55rem;
    padding:.7rem .78rem;
    border:1px solid #e2e8f0;
    border-radius:12px;
    background:#fff;
}
.pdf-download-row .format-note{margin:.12rem 0 .38rem}
.section-kicker{margin-top:.85rem;color:#0f172a;font-size:1.12rem;font-weight:850;letter-spacing:-.015em}
.section-kicker-copy{margin:.12rem 0 .55rem;color:#64748b;font-size:.72rem;line-height:1.35}
.st-key-regenerate_report button{
    min-height:2.2rem!important;
    padding:.32rem .72rem!important;
    border-radius:9px!important;
    border:1px solid #cbd5e1!important;
    background:#fff!important;
    color:#475569!important;
    font-size:.72rem!important;
    font-weight:700!important;
}
.st-key-html_report_download button::before{content:'↗';margin-right:.38rem;font-weight:900}
.st-key-pdf_report_download button::before{content:'↓';margin-right:.38rem;font-weight:900}
@media(max-width:620px){
    .report-complete-card{gap:.55rem;padding:.68rem .72rem;margin:.28rem 0 .58rem}
    .report-complete-icon{width:1.7rem;height:1.7rem;flex-basis:1.7rem;font-size:.86rem}
    .report-complete-title{font-size:.88rem}
    .report-complete-copy{font-size:.68rem}
    .download-hero{padding:.7rem;margin-bottom:.5rem;border-radius:12px}
    .download-format-title{font-size:.86rem}
    .download-feature-list{font-size:.62rem;gap:.22rem .48rem;margin:.34rem 0 .45rem}
    .pdf-download-row{padding:.58rem .62rem}
    .section-kicker{font-size:1rem;margin-top:.7rem}
    .section-kicker-copy{font-size:.66rem}
}


/* ================================================================
   Final product polish
   A unified spacing system, softer surfaces, restrained gradients,
   and clearer desktop/mobile action hierarchy.
   ================================================================ */

:root{
    --app-bg:#fbfcfe;
    --surface:#ffffff;
    --surface-soft:rgba(248,250,252,.88);
    --line-soft:rgba(148,163,184,.28);
    --ink:#172033;
    --muted:#64748b;
    --brand-red:#e5484d;
    --brand-red-deep:#d93f45;
    --brand-blue:#2f6feb;
    --brand-blue-deep:#2457d6;
    --radius-sm:10px;
    --radius-md:14px;
    --radius-lg:18px;
    --shadow-soft:0 8px 24px rgba(15,23,42,.055);
    --shadow-button:0 8px 20px rgba(15,23,42,.10);
}

html, body, [data-testid="stAppViewContainer"]{
    background:var(--app-bg)!important;
}

.block-container{
    max-width:1450px!important;
    padding-top:.55rem!important;
    padding-bottom:2.6rem!important;
}

/* Consistent section rhythm. */
h2, div[data-testid="stHeadingWithActionElements"] h2{
    margin-top:1.65rem!important;
    margin-bottom:.55rem!important;
}
h3{
    margin-top:1.15rem!important;
    margin-bottom:.42rem!important;
}
.stCaption, div[data-testid="stCaptionContainer"]{
    margin-bottom:.55rem!important;
}

/* Shared soft-surface language. */
div[data-testid="stVerticalBlockBorderWrapper"],
.validation-shell,
.executive-panel,
.download-hero,
.pdf-download-row,
.report-complete-card,
div[data-testid="stExpander"]{
    box-shadow:var(--shadow-soft)!important;
}

/* Generate action: compact, centered and visually related to the rest of
   the product rather than a full-width warning-red bar. */
.st-key-generate_report{
    display:flex!important;
    justify-content:center!important;
    margin:.35rem 0 .8rem!important;
}
.st-key-generate_report button{
    width:min(100%,460px)!important;
    min-height:48px!important;
    padding:.68rem 1.4rem!important;
    border:1px solid rgba(172,43,48,.22)!important;
    border-radius:13px!important;
    background:linear-gradient(180deg,rgba(239,87,92,.96),rgba(222,66,72,.96))!important;
    color:#fff!important;
    font-size:.92rem!important;
    font-weight:800!important;
    letter-spacing:.005em!important;
    box-shadow:0 9px 22px rgba(229,72,77,.20)!important;
    transition:transform .18s ease, box-shadow .18s ease, filter .18s ease!important;
}
.st-key-generate_report button:hover:not(:disabled){
    transform:translateY(-1px)!important;
    box-shadow:0 12px 26px rgba(229,72,77,.25)!important;
    filter:saturate(.96) brightness(1.02)!important;
}
.st-key-generate_report button:disabled{
    opacity:.48!important;
    box-shadow:none!important;
}

/* Regenerate remains intentionally quiet. */
.st-key-regenerate_report{
    display:flex!important;
    justify-content:flex-end!important;
    margin:.1rem 0 .35rem!important;
}
.st-key-regenerate_report button{
    background:rgba(255,255,255,.78)!important;
    backdrop-filter:blur(6px)!important;
    box-shadow:0 3px 10px rgba(15,23,42,.04)!important;
}

/* Softer, faded action buttons with restrained motion. */
.st-key-html_report_download button{
    min-height:48px!important;
    border:1px solid rgba(37,87,214,.22)!important;
    border-radius:13px!important;
    background:linear-gradient(180deg,rgba(58,117,239,.96),rgba(36,87,214,.96))!important;
    box-shadow:0 10px 24px rgba(47,111,235,.18)!important;
    transition:transform .18s ease, box-shadow .18s ease, filter .18s ease!important;
}
.st-key-html_report_download button:hover{
    transform:translateY(-1px)!important;
    box-shadow:0 13px 28px rgba(47,111,235,.23)!important;
    filter:brightness(1.025)!important;
}

.st-key-pdf_report_download button{
    min-height:46px!important;
    border:1px solid rgba(100,116,139,.34)!important;
    border-radius:13px!important;
    background:linear-gradient(180deg,rgba(255,255,255,.96),rgba(247,249,252,.96))!important;
    box-shadow:0 6px 16px rgba(15,23,42,.055)!important;
    transition:transform .18s ease, box-shadow .18s ease, background .18s ease!important;
}
.st-key-pdf_report_download button:hover{
    transform:translateY(-1px)!important;
    background:linear-gradient(180deg,#fff,#f1f5f9)!important;
    box-shadow:0 9px 20px rgba(15,23,42,.08)!important;
}

/* Completion/download surfaces use matching padding and radii. */
.report-complete-card{
    margin:.35rem 0 .65rem!important;
    padding:.72rem .82rem!important;
    border-radius:var(--radius-md)!important;
    background:linear-gradient(135deg,rgba(240,253,244,.92),rgba(248,255,251,.92) 55%,rgba(239,246,255,.90))!important;
}
.download-hero{
    margin:.15rem 0 .62rem!important;
    padding:.78rem .84rem .72rem!important;
    border-radius:var(--radius-md)!important;
    background:linear-gradient(145deg,rgba(248,251,255,.94),rgba(238,246,255,.90))!important;
}
.pdf-download-row{
    margin:.18rem 0 .58rem!important;
    padding:.72rem .82rem!important;
    border-radius:var(--radius-md)!important;
    background:linear-gradient(180deg,rgba(255,255,255,.96),rgba(250,251,253,.94))!important;
}
.download-feature-list{
    margin:.38rem 0 .52rem!important;
}
.pdf-features{
    margin-bottom:.48rem!important;
}
.pdf-features .download-feature::before{
    color:#64748b!important;
}

/* Advanced downloads should read as a tertiary utility. */
div[data-testid="stExpander"]{
    border-color:rgba(148,163,184,.32)!important;
    background:rgba(255,255,255,.80)!important;
    backdrop-filter:blur(7px);
}
div[data-testid="stExpander"] summary{
    min-height:44px!important;
    padding:.52rem .72rem!important;
}

/* Step indicators: softer and more cohesive. */
.process-strip{
    gap:.52rem!important;
    margin:.45rem 0 1.35rem!important;
}
.process-item{
    min-height:74px!important;
    display:flex!important;
    align-items:center!important;
    justify-content:center!important;
    border-color:rgba(148,163,184,.30)!important;
    background:linear-gradient(180deg,rgba(255,255,255,.93),rgba(249,251,253,.90))!important;
    box-shadow:0 5px 16px rgba(15,23,42,.035)!important;
    transition:transform .18s ease, box-shadow .18s ease!important;
}
.process-item:hover{
    transform:translateY(-1px);
    box-shadow:0 8px 20px rgba(15,23,42,.055)!important;
}
.process-item b{
    background:linear-gradient(135deg,#eef5ff,#e7f0ff)!important;
    box-shadow:inset 0 0 0 1px rgba(37,99,235,.06)!important;
}

/* Validation list: consistent compact spacing and softer success tint. */
.validation-shell{
    margin:.08rem 0 .72rem!important;
    padding:.52rem!important;
    border-radius:var(--radius-md)!important;
    background:linear-gradient(180deg,rgba(251,253,255,.94),rgba(248,250,252,.92))!important;
}
.validation-grid{
    gap:.34rem!important;
}
.validation-card{
    padding:.40rem .46rem!important;
    border-radius:9px!important;
}
.validation-card.detected{
    border-color:rgba(74,222,128,.44)!important;
    background:rgba(240,253,244,.78)!important;
}
.detected .validation-badge{
    background:rgba(220,252,231,.92)!important;
}

/* Executive highlight cards: tighter, softer and visually consistent. */
.executive-panel{
    margin-top:.25rem!important;
    padding:.62rem!important;
    border-radius:var(--radius-md)!important;
    background:linear-gradient(180deg,rgba(251,253,255,.95),rgba(247,250,252,.92))!important;
}
.executive-grid{
    gap:.55rem!important;
}
.executive-card{
    min-height:78px!important;
    padding:.62rem .68rem!important;
    border-color:rgba(148,163,184,.28)!important;
    background:rgba(255,255,255,.92)!important;
    box-shadow:0 5px 16px rgba(15,23,42,.035)!important;
    transition:transform .18s ease, box-shadow .18s ease!important;
}
.executive-card:hover{
    transform:translateY(-1px);
    box-shadow:0 8px 20px rgba(15,23,42,.055)!important;
}
.executive-card-value{
    margin-top:.12rem!important;
}
.executive-card-note{
    margin-top:.22rem!important;
}

/* Summary: more readable measure and restrained padding. */
.summary-card{
    max-width:1100px!important;
    margin:.62rem 0 0!important;
    padding:.66rem .78rem!important;
    font-size:.79rem!important;
    line-height:1.42!important;
    background:rgba(248,250,252,.90)!important;
}

/* Keep validation warning and details visually connected. */
div[data-testid="stAlert"] + div[data-testid="stExpander"]{
    margin-top:-.1rem!important;
}

/* Report tables get a quieter, product-style tab treatment. */
div[data-testid="stTabs"] [data-baseweb="tab-list"]{
    border:0!important;
    border-bottom:1px solid rgba(148,163,184,.28)!important;
    border-radius:0!important;
    background:transparent!important;
    box-shadow:none!important;
    padding:0!important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"]{
    min-height:40px!important;
    border-radius:9px 9px 0 0!important;
    background:transparent!important;
    box-shadow:none!important;
    transition:background .18s ease,color .18s ease!important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"]{
    color:var(--brand-red)!important;
    background:rgba(255,255,255,.66)!important;
    box-shadow:inset 0 -3px 0 var(--brand-red)!important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"]:first-child,
div[data-testid="stTabs"] button[data-baseweb="tab"]:first-child[aria-selected="true"]{
    background:transparent!important;
    color:inherit!important;
}
div[data-testid="stTabs"] button[data-baseweb="tab"]:first-child[aria-selected="true"]{
    color:var(--brand-red)!important;
    box-shadow:inset 0 -3px 0 var(--brand-red)!important;
}

/* Dataframes and upload surfaces use the same radius family. */
div[data-testid="stDataFrame"],
div[data-testid="stFileUploaderDropzone"]{
    border-radius:var(--radius-md)!important;
}

/* Desktop form and uploader spacing normalization. */
.st-key-report_details_card{
    padding:.52rem .66rem!important;
}
.st-key-report_details_card div[data-testid="stVerticalBlock"]{
    gap:.25rem!important;
}
.upload-guidance{
    margin:.18rem 0 .58rem!important;
}
div[data-testid="stFileUploader"]{
    margin-bottom:.2rem!important;
}

/* Compact, balanced report-format cards. Download actions no longer stretch
   into full-width bars on desktop. */
.st-key-html_format_card,
.st-key-pdf_format_card{
    min-height:174px;
    padding:.78rem .84rem!important;
    border:1px solid rgba(148,163,184,.26)!important;
    border-radius:14px!important;
    background:linear-gradient(180deg,rgba(255,255,255,.96),rgba(248,250,252,.92))!important;
    box-shadow:0 5px 16px rgba(15,23,42,.045)!important;
}
.st-key-html_format_card{
    border-color:rgba(37,99,235,.24)!important;
    background:linear-gradient(145deg,rgba(248,251,255,.96),rgba(239,246,255,.90))!important;
}
.st-key-html_report_download,
.st-key-pdf_report_download{
    display:flex!important;
    justify-content:flex-start!important;
}
.st-key-html_report_download button,
.st-key-pdf_report_download button{
    width:auto!important;
    min-width:180px!important;
    max-width:230px!important;
    padding-left:1rem!important;
    padding-right:1rem!important;
}

/* Mobile keeps clear hierarchy but avoids oversized blocks. */
@media(max-width:620px){
    .block-container{
        padding-left:.82rem!important;
        padding-right:.82rem!important;
        padding-bottom:1.55rem!important;
    }
    h2, div[data-testid="stHeadingWithActionElements"] h2{
        margin-top:1.15rem!important;
        margin-bottom:.34rem!important;
    }
    .st-key-generate_report button{
        width:100%!important;
        min-height:45px!important;
        font-size:.86rem!important;
    }
    .process-strip{
        gap:.28rem!important;
        margin:.25rem 0 .78rem!important;
    }
    .process-item{
        min-height:58px!important;
        padding:.36rem .16rem!important;
    }
    .report-complete-card{
        padding:.62rem .68rem!important;
    }
    .download-hero,
    .pdf-download-row{
        padding:.62rem .68rem!important;
    }
    .st-key-html_format_card,
    .st-key-pdf_format_card{
        min-height:0;
        padding:.68rem .72rem!important;
    }
    .st-key-html_report_download button,
    .st-key-pdf_report_download button{
        width:100%!important;
        max-width:none!important;
    }
    .executive-panel{
        padding:.46rem!important;
    }
    .executive-grid{
        gap:.38rem!important;
    }
    .executive-card{
        min-height:72px!important;
        padding:.52rem .56rem!important;
    }
    .summary-card{
        max-width:100%!important;
        font-size:.74rem!important;
        padding:.58rem .64rem!important;
    }
    .validation-shell{
        padding:.34rem!important;
    }
    .validation-card{
        padding:.28rem .34rem!important;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"]{
        padding:.42rem .58rem!important;
    }
}

</style>
""", unsafe_allow_html=True)

render_platform_hero()

st.markdown('<div class="page-nav-wrap"></div>', unsafe_allow_html=True)
generate_tab, preview_tab, guide_tab = st.tabs(
    ["Generate Report", "Example Report", "Guide"]
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
        <div class="about-panel"><strong>About this application</strong>This proof of concept transforms standardized Tableau exports into validated HTML, Excel and PDF reporting outputs. It is designed to reduce repetitive report preparation and give store teams more time for analysis and operational decisions.</div>
        """,
        unsafe_allow_html=True,
    )

with preview_tab:
    st.subheader("Example report")
    st.info("The example report uses synthetic sample data for demonstration. It does not contain actual store performance information.")
    st.caption("Explore representative sections from the generated report before uploading any files.")

    def example_visual(filename: str, alt_text: str) -> str:
        uri = image_data_uri(EXAMPLE_ASSET_DIR / filename)
        if uri:
            return f'<img class="demo-image" src="{uri}" alt="{html.escape(alt_text)}">'
        return '<div class="demo-placeholder">Example screenshot coming soon</div>'

    st.markdown(
        f"""
        <div class="demo-grid">
            <div class="demo-card">{example_visual('executive_dashboard.png', 'Executive KPI dashboard')}<div class="demo-title"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M4 20V10M10 20V4M16 20v-7M22 20V7"/></svg></span><strong>Executive Dashboard</strong></div><small>Headline KPIs, year-over-year movement, sales contribution and budget performance in one management view.</small></div>
            <div class="demo-card">{example_visual('department_analysis.png', 'Department and subdepartment analysis')}<div class="demo-title"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z"/></svg></span><strong>Department Analysis</strong></div><small>Sales, targets, budget gaps and contribution metrics organized by department and subdepartment.</small></div>
            <div class="demo-card">{example_visual('brands_products.png', 'Top brands and products')}<div class="demo-title"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M20 13 11 4H4v7l9 9 7-7Z"/><circle cx="7.5" cy="7.5" r="1"/></svg></span><strong>Brands &amp; Products</strong></div><small>Top-selling products by customer segment, with sales, units and low-stock visibility.</small></div>
            <div class="demo-card">{example_visual('executive_summary.png', 'Automated executive summary')}<div class="demo-title"><span class="mini-icon"><svg viewBox="0 0 24 24"><path d="M5 3h14v18H5z"/><path d="M8 8h8M8 12h8M8 16h5"/></svg></span><strong>Executive Summary</strong></div><small>Rule-based findings translate report results into concise performance highlights and operational priorities.</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    available_sample_outputs = [
        path for path in (SAMPLE_REPORT_PDF, SAMPLE_REPORT_HTML) if path.exists()
    ]
    if available_sample_outputs:
        st.markdown("#### Download the synthetic sample")
        sample_columns = st.columns(len(available_sample_outputs))
        for column, sample_path in zip(sample_columns, available_sample_outputs):
            with column:
                st.download_button(
                    f"Download sample {sample_path.suffix.lstrip('.').upper()}",
                    data=sample_path.read_bytes(),
                    file_name=sample_path.name,
                    mime="application/pdf" if sample_path.suffix.lower() == ".pdf" else "text/html",
                    use_container_width=True,
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
    """Embed generated chart images so HTML and PDF outputs are self-contained.

    Resolve the chart folder relative to this application file instead of the
    process working directory. Streamlit Cloud does not guarantee that the
    working directory used while generating a report matches the app folder.
    """
    chart_dir = ASSETS_DIR / "charts"
    if not chart_dir.exists():
        return html_text

    for image_path in chart_dir.glob("*.png"):
        encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
        data_uri = f"data:image/png;base64,{encoded}"
        html_text = html_text.replace(f"../assets/charts/{image_path.name}", data_uri)
        html_text = html_text.replace(f"assets/charts/{image_path.name}", data_uri)

    return html_text


def export_pdf(html_text: str) -> tuple[bytes, str | None]:
    """Convert the completed report HTML into a directly downloadable PDF.

    Returns ``(pdf_bytes, error_message)``. Keeping the error visible is
    important on Streamlit Cloud because missing Linux libraries or an
    incompatible WeasyPrint dependency otherwise look like a broken button.
    """
    try:
        from weasyprint import HTML

        pdf_bytes = HTML(string=html_text, base_url=str(APP_DIR)).write_pdf()
        if not isinstance(pdf_bytes, bytes):
            pdf_bytes = bytes(pdf_bytes)
        if len(pdf_bytes) < 5 or not pdf_bytes.startswith(b"%PDF"):
            raise RuntimeError("The PDF renderer returned an invalid document.")
        return pdf_bytes, None
    except Exception as exc:
        diagnostic = (
            f"{type(exc).__name__}: {exc}\n\n"
            "The HTML and Excel outputs were generated successfully. "
            "If this occurs on Streamlit Community Cloud, verify the "
            "WeasyPrint Python and Linux package dependencies, then reboot "
            "the deployment so they are installed during a fresh build."
        )
        return b"", diagnostic


@st.cache_data(show_spinner=False)
def build_outputs_from_uploads(
    file_bytes: Dict[str, bytes],
    store_name: str,
    report_period: str,
    template_version: str,
) -> Tuple[bytes, bytes, bytes, bytes, Dict[str, pd.DataFrame], Dict[str, pd.DataFrame | str]]:
    # template_version is intentionally part of the cached function arguments.
    # It prevents Streamlit from returning HTML/PDF created by an older report
    # template when the same uploads and report details are reused.
    del template_version
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

        pdf_bytes, pdf_error = export_pdf(html_text)
        outputs["pdf_error"] = pdf_error or ""

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
        st.warning(f"Validation warning: {passed_count}/{total_count} checks passed. Review differences before sharing the report.")
    with st.expander("View validation details", expanded=False):
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
        '''<div class="process-strip" aria-label="Report workflow">
            <div class="process-item"><b>1</b>Report details</div>
            <div class="process-item"><b>2</b>Upload exports</div>
            <div class="process-item"><b>3</b>Validate</div>
            <div class="process-item"><b>4</b>Generate &amp; download</div>
        </div>''',
        unsafe_allow_html=True,
    )
    st.subheader("Report details")
    st.caption("Add the location and reporting period that should appear in the report header.")

    with st.container(border=True, key="report_details_card"):
        store_name = st.text_input(
            "Store or location",
            value="",
            placeholder="e.g., Sample Store",
            help="Required. The Tableau exports do not reliably contain the store name.",
            key="report_store_name",
        )

        # A horizontal container is used instead of st.columns so the two short
        # date controls remain on one row on narrow screens in Streamlit 1.58.
        with st.container(
            horizontal=True,
            horizontal_alignment="distribute",
            vertical_alignment="top",
            gap="xsmall",
            key="report_date_row",
        ):
            report_start = st.date_input(
                "Report start date",
                value=date.today(),
                key="report_start_date",
                width="stretch",
            )
            report_end = st.date_input(
                "Report end date",
                value=date.today(),
                key="report_end_date",
                width="stretch",
            )

        default_report_period = f"{report_start.strftime('%b %d, %Y')} – {report_end.strftime('%b %d, %Y')}"
        with st.container(key="report_label_area"):
            report_period = st.text_input(
                "Report label (optional — shown in report)",
                value=default_report_period,
                help="This is the exact period label shown in the report header. Leave the generated date range as-is or edit it when a custom label is needed.",
                key="report_display_label",
            )
            st.caption("Uses the selected date range automatically unless you replace it with a custom reporting label.")
        report_period = report_period.strip() or default_report_period

    st.subheader("Upload Tableau exports")
    st.markdown(
        '<div class="upload-guidance"><strong>Eight exports are required.</strong> Select all files together. The platform identifies each report from its contents and checks that the complete set is present.<br><small>Files are processed within the application to build the report. Confirm your organization\'s data-handling requirements before uploading sensitive information.</small></div>',
        unsafe_allow_html=True,
    )

    bulk_uploads = st.file_uploader(
        "Upload all Tableau exports",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="bulk_tableau_uploads",
        help="You may select all eight files at once. Original filenames can vary because the app also checks their contents.",
    )

    validation = classify_uploaded_files(bulk_uploads or [])

    st.subheader("Validate exports")
    show_upload_validation(validation)

    current_fingerprint = upload_fingerprint(
        bulk_uploads or [], store_name, report_start, report_end, report_period
    )
    if st.session_state.get("input_fingerprint") != current_fingerprint:
        clear_generated_report()
        st.session_state["input_fingerprint"] = current_fingerprint

    location_missing = not store_name.strip()
    dates_invalid = report_start > report_end

    if location_missing:
        st.warning("Enter the store / location before generating the report.")
    if dates_invalid:
        st.warning("The report start date cannot be after the report end date.")
    ready_to_generate = bool(
        validation.get("ready")
        and not location_missing
        and not dates_invalid
    )

    report_is_ready = bool(st.session_state.get("report_ready"))
    st.subheader("Download report" if report_is_ready else "Generate report")
    if not ready_to_generate:
        st.warning("Missing required files or report details. Complete the items above to continue.")

    if report_is_ready:
        build_clicked = st.button(
            "Regenerate report",
            type="secondary",
            use_container_width=False,
            disabled=not ready_to_generate,
            key="regenerate_report",
        )
    else:
        build_clicked = st.button(
            "Generate Report",
            type="primary",
            use_container_width=False,
            disabled=not ready_to_generate,
            key="generate_report",
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
                    REPORT_TEMPLATE_VERSION,
                )

                st.session_state["model_bytes"] = model_bytes
                st.session_state["calculations_bytes"] = calculations_bytes
                st.session_state["html_report_bytes"] = html_report_bytes
                st.session_state["pdf_report_bytes"] = pdf_report_bytes
                st.session_state["tables"] = tables
                st.session_state["outputs"] = outputs
                st.session_state["report_ready"] = True

            except Exception as exc:
                clear_generated_report()
                st.error("The report could not be generated.")
                st.exception(exc)

    if st.session_state.get("report_ready"):
        outputs = st.session_state["outputs"]
        report_date_text = report_end.strftime("%Y-%m-%d")
        base_name = safe_filename(f"{store_name}_VM_KPI_Report_{report_date_text}")
        pdf_data = st.session_state.get("pdf_report_bytes", b"")
        pdf_error = str(st.session_state.get("outputs", {}).get("pdf_error", "")).strip()

        st.markdown(
            '<div class="report-complete-card">'
            '<div class="report-complete-icon">✓</div>'
            '<div><div class="report-complete-title">Report ready</div>'
            '<div class="report-complete-copy">All eight Tableau exports were validated and the report was generated successfully. Choose a format below or review the highlights.</div></div>'
            '</div>',
            unsafe_allow_html=True,
        )

        html_column, pdf_column = st.columns(2, gap="medium")

        with html_column:
            with st.container(border=True, key="html_format_card"):
                st.markdown(
                    '<div class="download-hero-top">'
                    '<div class="download-format-title"><span class="download-format-icon">🌐</span>Interactive HTML report</div>'
                    '<span class="recommended-tag">Recommended</span>'
                    '</div>'
                    '<div class="format-note">Standalone file for browser viewing, sharing and desktop access.</div>'
                    '<div class="download-feature-list">'
                    '<span class="download-feature">Browser ready</span>'
                    '<span class="download-feature">Interactive</span>'
                    '<span class="download-feature">Easy to share</span>'
                    '</div>',
                    unsafe_allow_html=True,
                )
                st.download_button(
                    "Download HTML",
                    data=st.session_state["html_report_bytes"],
                    file_name=f"{base_name}.html",
                    mime="text/html",
                    use_container_width=False,
                    key="html_report_download",
                )

        with pdf_column:
            with st.container(border=True, key="pdf_format_card"):
                st.markdown(
                    '<div class="download-format-title"><span class="download-format-icon">🖨️</span>Printable PDF report</div>'
                    '<div class="format-note">Fixed layout for printing, email and archiving.</div>'
                    '<div class="download-feature-list pdf-features">'
                    '<span class="download-feature">Print ready</span>'
                    '<span class="download-feature">Email friendly</span>'
                    '<span class="download-feature">Fixed layout</span>'
                    '</div>',
                    unsafe_allow_html=True,
                )
                if pdf_data:
                    st.download_button(
                        "Download PDF",
                        data=pdf_data,
                        file_name=f"{base_name}.pdf",
                        mime="application/pdf",
                        use_container_width=False,
                        key="pdf_report_download",
                    )
                elif pdf_error:
                    st.error("PDF unavailable; the HTML report is ready.")
                    with st.expander("PDF conversion details"):
                        st.code(pdf_error)
                else:
                    st.info("PDF export is temporarily unavailable.")

        with st.expander("Advanced downloads", expanded=False):
            st.caption("Technical workbooks for auditing, analysis and troubleshooting.")
            c1, c2 = st.columns(2, gap="small")
            with c1:
                st.download_button("Data model workbook", data=st.session_state["model_bytes"], file_name=f"{base_name}_Data_Model.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="data_model_download")
            with c2:
                st.download_button("Calculations workbook", data=st.session_state["calculations_bytes"], file_name=f"{base_name}_Calculations.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="calculations_download")

        st.markdown('<div class="section-kicker">Report highlights</div><div class="section-kicker-copy">A quick view of the most important results before opening the full report.</div>', unsafe_allow_html=True)
        show_kpi_cards(outputs["kpi_master"])
        summary_text = outputs.get("executive_summary_text", "")
        if summary_text:
            safe_summary = html.escape(str(summary_text)).replace("\n", "<br>")
            st.markdown(f'<div class="summary-card"><span class="summary-label">Auto executive summary</span>{safe_summary}</div>', unsafe_allow_html=True)

        st.subheader("Data validation checks")
        show_validation(outputs["validation"])
        st.subheader("Report tables")
        st.caption("Browse the detailed calculations and supporting tables used to build the report.")
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["Department KPIs", "Subdepartment Gaps", "Opportunities", "Top Brands", "Traffic"])
        with tab1:
            st.dataframe(outputs["department_kpis"], use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(outputs["sub_department_kpis"], use_container_width=True, hide_index=True)
        with tab3:
            st.caption("Priority Score = absolute budget gap % weighted by sales mix. It ranks areas that are both underperforming and meaningful to the business.")
            st.dataframe(outputs["opportunity_ranking"], use_container_width=True, hide_index=True)
        with tab4:
            st.dataframe(outputs["top_brands_fixed"], use_container_width=True, hide_index=True)
        with tab5:
            st.dataframe(outputs["traffic_summary"], use_container_width=True, hide_index=True)
            st.dataframe(outputs["hourly_traffic"], use_container_width=True, hide_index=True)
