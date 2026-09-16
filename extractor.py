"""Utilities for extracting NBTC EMF reports into a consistent table."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import re
import unicodedata
from typing import Iterable, Sequence

import pandas as pd
import pdfplumber
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


REPORT_COLUMNS = [
    "ลำดับที่",
    "ผู้ประกอบการ",
    "เลขที่ใบอนุญาตตั้ง",
    "ที่ตั้ง",
    "ตำบล",
    "อำเภอ",
    "จังหวัด",
    "รหัสไปรษณีย์",
    "Longitude",
    "Latitude",
    "ความถี่",
    "ตราอักษร",
    "รุ่น/แบบ",
    "กำลังส่ง (วัตต์)",
    "อัตราขยายสายอากาศ (dBi)",
    "ความสูงสายอากาศ (เมตร)",
    "ระยะห่างจากเสา ที่ตั้งสายอากาศ",
    "ระยะที่วัด/คำนวณ (เมตร)",
    "ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด",
    "วันที่วัด/คำนวณ",
    "ลงชื่อ",
    "วันที่รายงาน",
]

QA_COLUMNS = [
    "ไฟล์ต้นทาง",
    "หน้า",
    "สถานะ",
    "ข้อมูลที่ไม่พบ",
    "จำนวนความถี่",
    "หมายเหตุ",
]


# pdfplumber can preserve the visual spacing found in some Thai PDF fonts.
# These replacements are deliberately limited to terms used as field labels.
_THAI_TYPO_FIXES = {
    "ต าบล": "ตำบล",
    "ตา บล": "ตำบล",
    "อ าเภอ": "อำเภอ",
    "อา เภอ": "อำเภอ",
    "จา กดั": "จำกัด",
    "ที่ต้งั": "ที่ตั้ง",
    "ที่ ตั้ ง": "ที่ตั้ง",
    "หนา้ สา รวจ": "หน้าสำรวจ",
    "วนั ที่": "วันที่",
    "วันที่วดั": "วันที่วัด",
    "คา นวณ": "คำนวณ",
    "คำ นวณ": "คำนวณ",
    "ผมู้ ีอา นาจ": "ผู้มีอำนาจ",
    "ผมู้ ีอำ นำจ": "ผู้มีอำนาจ",
    "กระทา การ": "กระทำการ",
    "บริษทั": "บริษัท",
    "รหัสไปรษณยี ์": "รหัสไปรษณีย์",
}


LABELS = {
    "operator": ("หน่วยงาน",),
    "license": ("เลขที่ใบอนุญาตตั้ง", "เลขที่ใบอนุญาตต้งั"),
    "location": ("ที่ตั้ง", "ที่ต้งั"),
    "subdistrict": ("ตำบล", "ตาบล"),
    "district": ("อำเภอ", "อาเภอ"),
    "province": ("จังหวัด",),
    "zipcode": ("รหัสไปรษณีย์", "รหัสไปรษณยี"),
    "longitude": ("longitude", "longtitude"),
    "latitude": ("latitude",),
    "date_calc": ("วันที่วัด/คำนวณ", "วันที่วัด/คานวณ"),
    "signature": ("ผู้มีอำนาจลงนาม", "ผู้มีอำนาจลงนำม"),
    "date_report": ("วันที่รายงาน",),
}

MISSING_FIELD_NAMES = {
    "operator": "ผู้ประกอบการ",
    "license": "เลขที่ใบอนุญาตตั้ง",
    "location": "ที่ตั้ง",
    "subdistrict": "ตำบล",
    "district": "อำเภอ",
    "province": "จังหวัด",
    "zipcode": "รหัสไปรษณีย์",
    "longitude": "Longitude",
    "latitude": "Latitude",
    "frequencies": "ความถี่",
    "max_dist_val": "ระยะที่วัด/คำนวณ",
    "max_rad": "ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด",
}


@dataclass
class StationPage:
    """A parsed report page and the review information associated with it."""

    source_file: str
    page_number: int
    data: dict[str, str]
    frequencies: list[dict[str, str]]
    missing_fields: list[str]
    note: str = ""


def clean_text(value: object) -> str:
    """Return a readable, single-line value without changing factual content."""
    if value is None:
        return ""

    text = unicodedata.normalize("NFC", str(value))
    text = text.replace("\u00a0", " ").replace("\n", " ")
    for incorrect, corrected in _THAI_TYPO_FIXES.items():
        text = text.replace(incorrect, corrected)
    return re.sub(r"\s+", " ", text).strip()


def label_key(value: object) -> str:
    """Normalise labels only, so harmless font-spacing differences still match."""
    text = clean_text(value).lower().replace("longtitude", "longitude")
    text = re.sub(r"[\s:()\[\]{}=_\-/]", "", text)
    # Thai combining marks are frequently emitted in a different order by PDFs.
    text = re.sub(r"[\u0e31-\u0e3a\u0e47-\u0e4e]", "", text)
    return text


def matches_label(value: object, aliases: Sequence[str]) -> bool:
    key = label_key(value)
    return bool(key) and any(label_key(alias) in key for alias in aliases)


def is_same_label(value: object, aliases: Sequence[str]) -> bool:
    key = label_key(value)
    return bool(key) and any(key == label_key(alias) for alias in aliases)


def clean_rows(tables: Iterable[Sequence[Sequence[object]]]) -> list[list[str]]:
    rows: list[list[str]] = []
    for table in tables:
        for row in table or []:
            cleaned = [clean_text(cell) for cell in (row or [])]
            if any(cleaned):
                rows.append(cleaned)
    return rows


def value_after_label(rows: Sequence[Sequence[str]], aliases: Sequence[str]) -> str:
    """Read a label's value from the same row, with a one-row fallback."""
    for row_index, row in enumerate(rows):
        for column_index, cell in enumerate(row):
            if not matches_label(cell, aliases):
                continue

            for candidate in row[column_index + 1 :]:
                if candidate and not is_same_label(candidate, aliases):
                    return candidate

            # Some report variants put the value in the row directly below.
            if row_index + 1 < len(rows):
                next_row = rows[row_index + 1]
                candidates = next_row[column_index:] + next_row[:column_index]
                for candidate in candidates:
                    if candidate and not is_same_label(candidate, aliases):
                        return candidate
    return ""


def is_frequency(value: str) -> bool:
    return bool(re.fullmatch(r"\d{2,5}(?:\.\d+)?", value.strip()))


def extract_frequency_rows(rows: Sequence[Sequence[str]]) -> list[dict[str, str]]:
    """Extract the six frequency-table fields from either supplied table layout."""
    frequencies: list[dict[str, str]] = []
    seen: set[tuple[str, str, str, str, str, str]] = set()

    for row in rows:
        populated = [cell for cell in row if cell]
        if len(populated) < 6 or not is_frequency(populated[0]):
            continue

        values = tuple(populated[:6])
        if values in seen:
            continue
        seen.add(values)
        frequencies.append(
            {
                "frequency": values[0],
                "brand": values[1],
                "model": values[2],
                "power": values[3],
                "gain": values[4],
                "height": values[5],
            }
        )
    return frequencies


def is_numeric_value(value: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)?%?", value.strip()))


def extract_maximum(rows: Sequence[Sequence[str]]) -> tuple[str, str, str]:
    """Find the final summary row rather than the similarly named table heading."""
    label = "ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด"
    candidates: list[tuple[str, str]] = []

    for row in rows:
        has_summary_label = any(label_key(label) in label_key(cell) for cell in row)
        numeric_values = [cell for cell in row if is_numeric_value(cell)]
        if has_summary_label and len(numeric_values) >= 2:
            candidates.append((numeric_values[0], numeric_values[1]))

    if candidates:
        distance, radiation = candidates[-1]
        return "ระดับสูงสุด", distance, radiation
    return "", "", ""


def join_frequency_values(frequencies: Sequence[dict[str, str]], field: str) -> str:
    values = [frequency[field] for frequency in frequencies if frequency.get(field)]
    return ", ".join(dict.fromkeys(values))


def parse_page(page: pdfplumber.page.Page, source_file: str, page_number: int) -> StationPage:
    rows = clean_rows(page.extract_tables())
    data = {field: value_after_label(rows, aliases) for field, aliases in LABELS.items()}
    max_text, max_distance, max_radiation = extract_maximum(rows)
    data.update(
        {
            "max_dist_text": max_text,
            "max_dist_val": max_distance,
            "max_rad": max_radiation,
        }
    )
    frequencies = extract_frequency_rows(rows)

    note = ""
    if not rows:
        note = "ไม่พบตารางข้อมูลที่อ่านได้ในหน้านี้"

    checks = {**data, "frequencies": "yes" if frequencies else ""}
    missing_fields = [
        display_name
        for field, display_name in MISSING_FIELD_NAMES.items()
        if not checks.get(field, "")
    ]
    return StationPage(source_file, page_number, data, frequencies, missing_fields, note)


def parse_pdf_bytes(source_file: str, source_bytes: bytes) -> tuple[list[StationPage], list[str]]:
    """Parse each report page independently so multi-station PDFs remain separated."""
    stations: list[StationPage] = []
    errors: list[str] = []

    try:
        with pdfplumber.open(BytesIO(source_bytes)) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                try:
                    station = parse_page(page, source_file, page_number)
                    if any(station.data.values()) or station.frequencies:
                        stations.append(station)
                    else:
                        errors.append(f"{source_file} หน้า {page_number}: ไม่พบข้อมูลสถานี")
                except Exception as error:  # Keep processing the remaining report pages.
                    errors.append(f"{source_file} หน้า {page_number}: {error}")
    except Exception as error:
        errors.append(f"เปิดไฟล์ {source_file} ไม่สำเร็จ: {error}")
    return stations, errors


def page_to_summary_record(station: StationPage) -> dict[str, str]:
    data = station.data
    return {
        "ผู้ประกอบการ": data["operator"],
        "เลขที่ใบอนุญาตตั้ง": data["license"],
        "ที่ตั้ง": data["location"],
        "ตำบล": data["subdistrict"],
        "อำเภอ": data["district"],
        "จังหวัด": data["province"],
        "รหัสไปรษณีย์": data["zipcode"],
        "Longitude": data["longitude"],
        "Latitude": data["latitude"],
        "ความถี่": join_frequency_values(station.frequencies, "frequency"),
        "ตราอักษร": join_frequency_values(station.frequencies, "brand"),
        "รุ่น/แบบ": join_frequency_values(station.frequencies, "model"),
        "กำลังส่ง (วัตต์)": join_frequency_values(station.frequencies, "power"),
        "อัตราขยายสายอากาศ (dBi)": join_frequency_values(station.frequencies, "gain"),
        "ความสูงสายอากาศ (เมตร)": join_frequency_values(station.frequencies, "height"),
        "ระยะห่างจากเสา ที่ตั้งสายอากาศ": data["max_dist_text"],
        "ระยะที่วัด/คำนวณ (เมตร)": data["max_dist_val"],
        "ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด": data["max_rad"],
        "วันที่วัด/คำนวณ": data["date_calc"],
        "ลงชื่อ": data["signature"],
        "วันที่รายงาน": data["date_report"],
    }


def station_to_detail_records(station: StationPage) -> list[dict[str, str]]:
    base = page_to_summary_record(station)
    if not station.frequencies:
        return [base]

    records: list[dict[str, str]] = []
    for frequency in station.frequencies:
        record = base.copy()
        record.update(
            {
                "ความถี่": frequency["frequency"],
                "ตราอักษร": frequency["brand"],
                "รุ่น/แบบ": frequency["model"],
                "กำลังส่ง (วัตต์)": frequency["power"],
                "อัตราขยายสายอากาศ (dBi)": frequency["gain"],
                "ความสูงสายอากาศ (เมตร)": frequency["height"],
            }
        )
        records.append(record)
    return records


def build_report_records(stations: Sequence[StationPage], detail_mode: bool) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for station in stations:
        page_records = station_to_detail_records(station) if detail_mode else [page_to_summary_record(station)]
        records.extend(page_records)

    for index, record in enumerate(records, start=1):
        record["ลำดับที่"] = index
    return records


def build_qa_records(stations: Sequence[StationPage]) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for station in stations:
        status = "พร้อมใช้งาน" if not station.missing_fields else "ต้องตรวจสอบ"
        records.append(
            {
                "ไฟล์ต้นทาง": station.source_file,
                "หน้า": station.page_number,
                "สถานะ": status,
                "ข้อมูลที่ไม่พบ": ", ".join(station.missing_fields),
                "จำนวนความถี่": len(station.frequencies),
                "หมายเหตุ": station.note,
            }
        )
    return records


def _format_worksheet(worksheet, *, status_column: int | None = None) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    warning_fill = PatternFill("solid", fgColor="FFF2CC")
    normal_fill = PatternFill("solid", fgColor="E2F0D9")

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.row_dimensions[1].height = 34

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    if status_column:
        for row_number in range(2, worksheet.max_row + 1):
            status_cell = worksheet.cell(row=row_number, column=status_column)
            status_cell.fill = normal_fill if status_cell.value == "พร้อมใช้งาน" else warning_fill

    for column_number, cells in enumerate(worksheet.iter_cols(), start=1):
        longest = max((len(str(cell.value or "")) for cell in cells), default=10)
        width = max(12, min(longest + 2, 42))
        worksheet.column_dimensions[get_column_letter(column_number)].width = width


def create_excel_bytes(
    report_records: Sequence[dict[str, str]], qa_records: Sequence[dict[str, str]]
) -> bytes:
    """Create a filterable report sheet plus a separate, non-blocking QA sheet."""
    report_df = pd.DataFrame(report_records, columns=REPORT_COLUMNS)
    qa_df = pd.DataFrame(qa_records, columns=QA_COLUMNS)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report_df.to_excel(writer, index=False, sheet_name="Report_Data")
        qa_df.to_excel(writer, index=False, sheet_name="ตรวจสอบข้อมูล")
        _format_worksheet(writer.sheets["Report_Data"])
        _format_worksheet(writer.sheets["ตรวจสอบข้อมูล"], status_column=3)

    return output.getvalue()
