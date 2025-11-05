# automation/multi_sheet_runner.py
from __future__ import annotations

import argparse

import json
import logging
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Optional

import pandas as pd

SHEET_PATTERN_DEFAULT = r"^cdas_\d+$"
BILL_COLUMNS = ("invoice", "bill", "bill_id")
DATE_COLUMNS = ("date_to_download", "download_date", "date")


def parse_args() -> argparse.Namespace:
    load_env_file()

    default_dir = os.environ.get("BILL_DATA_SRC_DIRECTORY")
    if not default_dir:
        username = os.environ.get("CDAS_USERNAME") or Path.home().name
        default_dir = str(Path(f"C:/Users/{username}/Downloads"))
    
    parser = argparse.ArgumentParser(
        description="Iterate Excel worksheets and trigger app.py Selenium automation."
    )
    parser.add_argument(
        "--workbook-dir",
        default=os.environ.get("BILL_DATA_SRC_DIRECTORY", "C:/Users/username/Downloads"),
        help="Directory containing the workbook; defaults to BILL_DATA_SRC_DIRECTORY or C:/Users/username/Downloads.",
    )
    parser.add_argument(
        "--workbook-name",
        default=os.environ.get("BILL_DATA_SRC", "cdas_n.xlsx"),
        help="Workbook filename; defaults to BILL_DATA_SRC or cdas_n.xlsx.",
    )
    parser.add_argument(
        "--sheet-pattern",
        default=SHEET_PATTERN_DEFAULT,
        help="Regex used to select worksheet names (default: ^cdas_\\d+$).",
    )
    parser.add_argument(
        "--app-path",
        default="app.py",
        help="Path to the Selenium automation entry point.",
    )
    parser.add_argument(
        "--log-file",
        default="logs/bill_download.log",
        help="File that receives structured run information.",
    )
    parser.add_argument(
        "--python",
        default=sys.executable,
        help="Python interpreter used to execute app.py.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Inspect workbook and log payloads without calling app.py.",
    )
    return parser.parse_args()

def load_env_file() -> None:
    env_path = Path(__file__).resolve().parent.parent / ".env"
    if env_path.exists():
        for raw in env_path.read_text().splitlines():
            raw = raw.strip()
            if not raw or raw.startswith("#"):
                continue
            key, _, value = raw.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))

user_from_env = os.environ.get("BILL_DATA_SRC_DIRECTORY")
if not user_from_env:
    username = os.environ.get("CDAS_USERNAME") or os.environ.get("USERNAME")
    downloads = Path.home() / "Downloads"
    if username:
        downloads = Path(f"C:/Users/{username}/Downloads")
    user_from_env = str(downloads)



def setup_logger(log_file: Path) -> logging.Logger:
    logger = logging.getLogger("bill-runner")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    log_file.parent.mkdir(parents=True, exist_ok=True)
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(stream_handler)

    return logger


def select_sheets(workbook: Path, pattern: re.Pattern[str]) -> list[str]:
    with pd.ExcelFile(workbook, engine="openpyxl") as excel:
        return [sheet for sheet in excel.sheet_names if pattern.search(sheet)]


def extract_bills(df: pd.DataFrame, sheet: str, workbook: Path) -> tuple[list[str], str]:
    for column in BILL_COLUMNS:
        if column in df.columns:
            bills = [
                value
                for value in df[column]
                .dropna()
                .astype(str)
                .str.strip()
                if value
            ]
            return bills, column
    raise KeyError(
        f"No bill/invoice column found in worksheet '{sheet}' from '{workbook}'. "
        f"Expected one of {BILL_COLUMNS}."
    )


def extract_date(df: pd.DataFrame) -> Optional[str]:
    for column in DATE_COLUMNS:
        if column in df.columns:
            series = df[column].dropna()
            if series.empty:
                continue
            try:
                parsed = pd.to_datetime(series.iloc[0], errors="coerce")
                if pd.isna(parsed):
                    continue
                return parsed.strftime("%Y-%m-%d")
            except Exception:
                return str(series.iloc[0])
    return None


def run_sheet(
    sheet: str,
    workbook: Path,
    app_path: Path,
    logger: logging.Logger,
    python_exec: str,
    dry_run: bool,
) -> bool:
    df = pd.read_excel(
        workbook,
        sheet_name=sheet,
        header=0,
        engine="openpyxl",
    )
    bills, bill_column = extract_bills(df, sheet, workbook)
    if not bills:
        logger.warning("Worksheet '%s' has no bill identifiers; skipping.", sheet)
        return True

    payload = {
        "worksheet": sheet,
        "date_to_download": extract_date(df),
        "bill_downloaded": bills,
        "len(bill_downloaded)": len(bills),
    }

    if dry_run:
        logger.info("Dry run for worksheet '%s': %s", sheet, json.dumps(payload))
        return True

    env = os.environ.copy()
    env["ESKER_INVOICE_WORKBOOK"] = str(workbook)
    env["ESKER_INVOICE_SHEET"] = sheet
    env["ESKER_INVOICE_COLUMN"] = bill_column

    logger.info("Launching automation for worksheet '%s' (%d bills).", sheet, len(bills))
    result = subprocess.run(
        [python_exec, str(app_path)],
        env=env,
        capture_output=True,
        text=True,
    )

    if result.stdout:
        logger.debug(result.stdout.strip())
    if result.stderr:
        logger.debug(result.stderr.strip())

    if result.returncode != 0:
        payload["error"] = result.stderr or result.stdout or "Unknown failure"
        logger.error("Automation failed for worksheet '%s'.", sheet)
        logger.error(json.dumps(payload))
        return False

    logger.info(json.dumps(payload))
    return True


def main() -> None:
    args = parse_args()
    workbook = (Path(args.workbook_dir).expanduser() / args.workbook_name).resolve()
    if not workbook.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook}")

    app_path = Path(args.app_path).resolve()
    if not app_path.exists():
        raise FileNotFoundError(f"app.py not found at {app_path}")

    logger = setup_logger(Path(args.log_file))
    sheet_pattern = re.compile(args.sheet_pattern, re.IGNORECASE)
    sheets = select_sheets(workbook, sheet_pattern)
    if not sheets:
        logger.warning(
            "No worksheet in '%s' matched pattern '%s'.",
            workbook,
            sheet_pattern.pattern,
        )
        sys.exit(1)

    failures = 0
    for sheet in sheets:
        if not run_sheet(sheet, workbook, app_path, logger, args.python, args.dry_run):
            failures += 1

    if failures:
        logger.warning(
            "Completed with %d failure(s) across %d worksheet(s).",
            failures,
            len(sheets),
        )
        sys.exit(1)

    logger.info("Automation finished successfully for %d worksheet(s).", len(sheets))


if __name__ == "__main__":
    main()

