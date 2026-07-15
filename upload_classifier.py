"""Automatic classifier and validator for VM KPI Tableau exports.

The Streamlit app accepts all source files in one bulk upload. This module:
- reads CSV and Excel uploads safely from memory;
- identifies each report using filename, columns, and sample values;
- reports missing, duplicate, unrecognized, and unreadable files;
- returns the original file bytes under the internal keys expected by the data model.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from io import BytesIO
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Mapping, Sequence

import pandas as pd


REPORT_DEFINITIONS: Dict[str, Dict[str, str]] = {
    "lfl": {"label": "LFL", "expected_file": "LFL.csv"},
    "group_by_site": {"label": "Group TY / Department Summary", "expected_file": "Group by Site.xlsx"},
    "d_net_sales": {"label": "D Net Sales / Department Detail", "expected_file": "D Net Sales.csv"},
    "footwear": {"label": "Footwear", "expected_file": "Footwear.csv"},
    "apparel": {"label": "Apparel", "expected_file": "Apparel.csv"},
    "accessories": {"label": "Accessories", "expected_file": "Accessories.csv"},
    "traffic_by_hour": {"label": "Traffic By The Hour / Table", "expected_file": "Table.csv"},
    "option_list": {"label": "Option List / Product Sales", "expected_file": "Option List.csv"},
}

REQUIRED_REPORT_KEYS: tuple[str, ...] = tuple(REPORT_DEFINITIONS)


@dataclass(frozen=True)
class ClassificationResult:
    filename: str
    report_type: str | None
    confidence: float
    reason: str
    error: str | None = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def _normalise_text(value: Any) -> str:
    text = str(value or "").replace("\xa0", " ").strip().lower()
    return re.sub(r"\s+", " ", text)


def _normalise_token(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", _normalise_text(value)).strip()


def _read_csv_bytes(data: bytes) -> pd.DataFrame:
    last_error: Exception | None = None
    for encoding in ("utf-16", "utf-8-sig", "utf-8", "latin1"):
        for sep in ("\t", ",", ";"):
            try:
                df = pd.read_csv(BytesIO(data), encoding=encoding, sep=sep)
                if len(df.columns) > 1:
                    return df
            except Exception as exc:  # pragma: no cover - retained for diagnostics
                last_error = exc
    raise ValueError(f"Could not parse the file as a Tableau CSV. Last error: {last_error}")


def read_uploaded_table(uploaded_file: Any) -> pd.DataFrame:
    """Read a Streamlit UploadedFile or a bytes-like object without consuming it."""
    name = getattr(uploaded_file, "name", "uploaded_file")
    if hasattr(uploaded_file, "getvalue"):
        data = uploaded_file.getvalue()
    elif isinstance(uploaded_file, (bytes, bytearray)):
        data = bytes(uploaded_file)
    else:
        raise TypeError("Unsupported uploaded file object.")

    suffix = Path(name).suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        try:
            return pd.read_excel(BytesIO(data), engine="openpyxl")
        except Exception:
            # Tableau files are sometimes renamed while still being delimited text.
            pass
    return _read_csv_bytes(data)


def _column_tokens(df: pd.DataFrame) -> set[str]:
    return {_normalise_token(col) for col in df.columns}


def _sample_text(df: pd.DataFrame, max_rows: int = 40, max_cols: int = 16) -> str:
    sample = df.iloc[:max_rows, :max_cols].fillna("").astype(str)
    values = [_normalise_text(v) for v in sample.to_numpy().ravel()]
    return " | ".join(values)


def _contains_any(text: str, terms: Iterable[str]) -> bool:
    return any(_normalise_text(term) in text for term in terms)


def _count_columns(columns: set[str], terms: Sequence[str]) -> int:
    normalised = {_normalise_token(term) for term in terms}
    return len(columns.intersection(normalised))


def identify_uploaded_file(uploaded_file: Any) -> ClassificationResult:
    """Identify one report using filename, columns, and representative cell values."""
    filename = getattr(uploaded_file, "name", "uploaded_file")
    filename_token = _normalise_token(Path(filename).stem)

    try:
        df = read_uploaded_table(uploaded_file)
    except Exception as exc:
        return ClassificationResult(filename, None, 0.0, "File could not be read.", str(exc))

    columns = _column_tokens(df)
    sample = _sample_text(df)
    full_signal = f"{filename_token} | {' | '.join(sorted(columns))} | {sample}"

    # Option List is structurally unique and should be checked first.
    option_score = _count_columns(columns, [
        "Style Id", "Style No", "Department Desc", "Brand Desc", "Style Desc",
        "Units Sold", "Units on Hand",
    ])
    if option_score >= 5:
        return ClassificationResult(filename, "option_list", 1.0, "Matched product/style and inventory columns.")

    # Traffic has a distinctive wide Tableau structure with hour labels and metric rows.
    hour_pattern = bool(re.search(r"\b\d{1,2}\s*[ap]\.?\s*m\.?\b", sample.replace(" ", "")))
    traffic_terms = sum(term in sample for term in ["footfall", "transactions", "conversion", "atv", "net sales"])
    if (hour_pattern and traffic_terms >= 2) or (
        _contains_any(filename_token, ["traffic", "table"]) and traffic_terms >= 2
    ):
        return ClassificationResult(filename, "traffic_by_hour", 0.98, "Matched hourly traffic labels and traffic metrics.")

    # LFL contains previous-year comparisons and IPC/UPT measures.
    lfl_terms = sum(term in full_signal for term in [
        "previous year net sales", "sales lfl", "transactions lfl", "previous year transactions",
        "upt lfl", "ipc lfl", "ly upt", "net sales previous year var",
    ])
    if lfl_terms >= 2 or ("lfl" in filename_token and lfl_terms >= 1):
        return ClassificationResult(filename, "lfl", 0.97, "Matched LFL and previous-year KPI fields.")

    # Department summary: one row per Apparel/Footwear/Accessories plus Avg Sales Price.
    summary_columns = _count_columns(columns, [
        "Group", "Net Sales", "Budget CA", "VAR (CA)", "VAR % (CA)", "Avg Sales Price", "Units",
    ])
    core_groups = sum(group in sample for group in ["footwear", "apparel", "accessories"])
    if summary_columns >= 4 and core_groups >= 2 and "avg sales price" in full_signal:
        return ClassificationResult(filename, "group_by_site", 0.96, "Matched department summary KPIs and core department rows.")

    # D Net Sales: subdepartment detail plus budget fields.
    detail_columns = _count_columns(columns, [
        "Department Description", "Budget CA", "VAR (CA)", "VAR % (CA)", "Net Sales", "Units",
    ])
    subdept_terms = sum(term in sample for term in [
        "mens footwear", "womens footwear", "junior footwear", "mens apparel",
        "womens apparel", "clothing accessories", "other accessories",
    ])
    if detail_columns >= 4 and subdept_terms >= 1:
        return ClassificationResult(filename, "d_net_sales", 0.96, "Matched subdepartment sales and budget fields.")

    # Brand exports share a schema; use department values and filename as secondary evidence.
    brand_columns = _count_columns(columns, [
        "Net Sales", "% of Net Sales", "% of Gender", "Margin", "Margin %", "Units Sold", "% of Units Sold",
    ])
    is_brand_structure = brand_columns >= 4 or (
        "units sold" in full_signal and "% of gender" in full_signal and "net sales" in full_signal
    )
    if is_brand_structure:
        category_signals = {
            "footwear": sum(term in sample for term in ["mens footwear", "womens footwear", "junior footwear", "childrens footwear", "infants footwear"]),
            "apparel": sum(term in sample for term in ["mens apparel", "womens apparel", "junior apparel", "childrens apparel", "infants apparel"]),
            "accessories": sum(term in sample for term in ["clothing accessories", "other accessories"]),
        }
        for key in category_signals:
            if key in filename_token:
                category_signals[key] += 2
        best_key = max(category_signals, key=category_signals.get)
        best_score = category_signals[best_key]
        tied = list(category_signals.values()).count(best_score) > 1
        if best_score > 0 and not tied:
            return ClassificationResult(filename, best_key, 0.94, f"Matched {best_key} brand-report structure and department values.")

    # Conservative filename fallbacks only when the file also contains plausible fields.
    filename_fallbacks = {
        "group by site": "group_by_site",
        "group ty": "group_by_site",
        "d net sales": "d_net_sales",
        "option list": "option_list",
        "footwear": "footwear",
        "apparel": "apparel",
        "accessories": "accessories",
        "traffic": "traffic_by_hour",
        "lfl": "lfl",
    }
    for phrase, report_type in filename_fallbacks.items():
        if phrase in filename_token and len(columns) >= 2:
            return ClassificationResult(filename, report_type, 0.70, f"Identified from filename with a valid tabular structure ({phrase}).")

    return ClassificationResult(filename, None, 0.0, "The file structure did not match a required Tableau export.")


def classify_uploaded_files(uploaded_files: Sequence[Any]) -> Dict[str, Any]:
    """Classify and validate all uploaded files.

    Extra unrecognized files do not block generation. Missing reports, duplicate
    report types, and unreadable files do block generation.
    """
    results = [identify_uploaded_file(file) for file in uploaded_files]
    grouped: Dict[str, List[Any]] = {key: [] for key in REQUIRED_REPORT_KEYS}
    result_by_name = {result.filename: result for result in results}

    for file, result in zip(uploaded_files, results):
        if result.report_type in grouped and result.error is None:
            grouped[result.report_type].append(file)

    classified = {key: files[0] for key, files in grouped.items() if len(files) == 1}
    duplicates = {
        key: [getattr(file, "name", "uploaded_file") for file in files]
        for key, files in grouped.items() if len(files) > 1
    }
    missing = [key for key in REQUIRED_REPORT_KEYS if key not in classified and key not in duplicates]
    unrecognized = [result.filename for result in results if result.report_type is None and result.error is None]
    errors = [result.to_dict() for result in results if result.error is not None]

    file_bytes = {
        key: file.getvalue() if hasattr(file, "getvalue") else bytes(file)
        for key, file in classified.items()
    }

    return {
        "classified": classified,
        "file_bytes": file_bytes,
        "missing": missing,
        "duplicates": duplicates,
        "unrecognized": unrecognized,
        "errors": errors,
        "results": [result.to_dict() for result in results],
        "ready": not missing and not duplicates and not errors,
        "result_by_name": result_by_name,
    }


def validation_rows(validation: Mapping[str, Any]) -> List[Dict[str, str]]:
    """Create presentation-ready rows for the Streamlit validation panel."""
    duplicate_keys = set(validation.get("duplicates", {}))
    missing_keys = set(validation.get("missing", []))
    rows: List[Dict[str, str]] = []

    for key, definition in REPORT_DEFINITIONS.items():
        if key in duplicate_keys:
            status = "Duplicate"
            files = ", ".join(validation["duplicates"][key])
        elif key in missing_keys:
            status = "Missing"
            files = "—"
        else:
            status = "Detected"
            file = validation.get("classified", {}).get(key)
            files = getattr(file, "name", "—")
        rows.append({"Status": status, "Required report": definition["label"], "Uploaded file": files})

    return rows
