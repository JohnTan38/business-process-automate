"""
Utility script to export Esker master data Excel sheets into plain-text payloads.

The script reads `master_gl.xlsx` and `master_vendor.xlsx` from the download
directory, shapes the important columns, and writes space-delimited text files
that mirror the automation email format.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Tuple

import unicodedata

import pandas as pd

MASTER_SRC_DIR = Path(r"C:/Users/john.tan/Downloads")
GL_OUTPUT_NAME = "master_gl.txt"
VENDOR_OUTPUT_NAME = "master_vendor.txt"


def _normalize_strings(series: pd.Series) -> pd.Series:
    """Return a series with whitespace-trimmed strings and empty values as ''."""
    return (
        series.fillna("")
        .astype(str)
        .map(lambda value: " ".join(str(value).split()).strip())
    )


def _unique_tokens(values: Iterable[str]) -> list[str]:
    """Return a list of unique tokens preserving the original order."""
    seen: set[str] = set()
    ordered: list[str] = []
    for item in values:
        token = str(item).strip()
        if not token or token in seen:
            continue
        seen.add(token)
        ordered.append(token)
    return ordered


def _clean_vendor_name(name: str) -> str:
    """Return ASCII vendor name without placeholder characters like '?'."""
    normalized = unicodedata.normalize("NFKD", name or "")
    filtered = "".join(
        ch for ch in normalized
        if not unicodedata.category(ch).startswith("C")
    )
    ascii_only = filtered.encode("ascii", errors="ignore").decode("ascii")
    cleaned = ascii_only.replace("?", "")
    return " ".join(cleaned.split())


def export_master_text(master_src_dir: Path | str = MASTER_SRC_DIR) -> Tuple[Path, Path]:
    """
    Read the master Excel sources and export text payloads.

    Returns:
        A tuple containing the paths to the GL and vendor text outputs.
    """
    base_dir = Path(master_src_dir)
    gl_excel = base_dir / "master_gl.xlsx"
    vendor_excel = base_dir / "master_vendor.xlsx"

    if not gl_excel.exists():
        raise FileNotFoundError(f"Missing GL master file: {gl_excel}")
    if not vendor_excel.exists():
        raise FileNotFoundError(f"Missing vendor master file: {vendor_excel}")

    gl_cols = ["account", "coding_block", "company_code", "description"]
    vendor_cols = ["company_code", "vendor_number", "vendor_name"]

    gl_df = pd.read_excel(gl_excel, dtype={col: str for col in gl_cols})
    vendor_df = pd.read_excel(vendor_excel, dtype={col: str for col in vendor_cols})

    missing_gl = [col for col in gl_cols if col not in gl_df.columns]
    if missing_gl:
        raise ValueError(f"GL master missing columns: {', '.join(missing_gl)}")
    missing_vendor = [col for col in vendor_cols if col not in vendor_df.columns]
    if missing_vendor:
        raise ValueError(f"Vendor master missing columns: {', '.join(missing_vendor)}")

    gl_df = gl_df[gl_cols].copy()
    vendor_df = vendor_df[vendor_cols].copy()

    for column in gl_cols:
        gl_df[column] = _normalize_strings(gl_df[column])
    for column in vendor_cols:
        vendor_df[column] = _normalize_strings(vendor_df[column])

    gl_df = gl_df[gl_df["account"].astype(bool) & gl_df["company_code"].astype(bool) & gl_df["description"].astype(bool)]
    vendor_df = vendor_df[vendor_df["company_code"].astype(bool) & vendor_df["vendor_number"].astype(bool) & vendor_df["vendor_name"].astype(bool)]

    gl_lines: list[str] = []
    for account, coding_block, company_code, description in gl_df.itertuples(index=False):
        tokens = [account]
        if coding_block:
            tokens.append(coding_block)
        tokens.extend(_unique_tokens(company_code.split(";")))
        line = " ".join(tokens + [description])
        gl_lines.append(line)

    vendor_lines: list[str] = []
    for company_code, vendor_number, vendor_name in vendor_df.itertuples(index=False):
        codes = _unique_tokens(company_code.split(";"))
        for code in codes:
            sanitized_name = _clean_vendor_name(vendor_name)
            if not sanitized_name:
                sanitized_name = "UNKNOWN_VENDOR"
            vendor_lines.append(f"{code} {vendor_number} {sanitized_name}")

    gl_output_path = base_dir / GL_OUTPUT_NAME
    vendor_output_path = base_dir / VENDOR_OUTPUT_NAME

    gl_output_path.write_text("\n".join(gl_lines), encoding="utf-8")
    vendor_output_path.write_text("\n".join(vendor_lines), encoding="utf-8")

    return gl_output_path, vendor_output_path


if __name__ == "__main__":  # pragma: no cover
    gl_path, vendor_path = export_master_text()
    print(f"GL master exported to: {gl_path}")
    print(f"Vendor master exported to: {vendor_path}")
