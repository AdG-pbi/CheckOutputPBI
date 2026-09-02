from __future__ import annotations

import argparse
import csv
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterator, Sequence

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from pypdf import PdfReader

TABULAR_EXTENSIONS = {".csv", ".xlsx", ".xlsm"}
DEFAULT_TABULAR_SECTION = "__tabular__"
MAX_EXCEL_ROWS = 1_048_576
STATUS_FILLS = {
    "ADDED": PatternFill(fill_type="solid", fgColor="C6EFCE"),
    "REMOVED": PatternFill(fill_type="solid", fgColor="FFC7CE"),
    "CHANGED": PatternFill(fill_type="solid", fgColor="FFEB9C"),
}


@dataclass
class ComparisonResult:
    mode: str
    differences: list[dict[str, str | int]]
    summary: dict[str, str | int]


@dataclass(frozen=True)
class TextEntry:
    text: str
    location: str


class ComparisonError(ValueError):
    pass


def _shorten_text(text: str, max_length: int = 60) -> str:
    collapsed = " ".join(text.split())
    if len(collapsed) <= max_length:
        return collapsed
    return f"{collapsed[: max_length - 3].rstrip()}..."


def build_entry_context(entries: Sequence[TextEntry], index: int) -> str:
    parts: list[str] = []
    if index > 0:
        parts.append(f"Dopo: {_shorten_text(entries[index - 1].text)}")
    if index + 1 < len(entries):
        parts.append(f"Prima di: {_shorten_text(entries[index + 1].text)}")
    return " | ".join(parts)


def parse_key_spec(key_spec: str | None) -> tuple[int, ...] | None:
    if not key_spec:
        return None

    indexes: list[int] = []
    for chunk in key_spec.split("+"):
        chunk = chunk.strip()
        if not chunk or not chunk.isdigit():
            raise ComparisonError(
                "La chiave deve essere nel formato '1+5' con indici numerici positivi."
            )
        index = int(chunk)
        if index <= 0:
            raise ComparisonError("Gli indici della chiave devono partire da 1.")
        indexes.append(index - 1)
    return tuple(indexes)


def read_csv_rows(path: Path) -> list[list[str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [[str(cell).strip() for cell in row] for row in csv.reader(handle)]


def _normalize_cell(cell: object) -> str:
    if cell is None:
        return ""
    if isinstance(cell, float) and cell.is_integer():
        return str(int(cell))
    return str(cell).strip()


def read_xlsx_rows(path: Path) -> list[list[str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    active_sheet = workbook.active
    rows = []
    for row in active_sheet.iter_rows(values_only=True):
        rows.append([_normalize_cell(cell) for cell in row])
    workbook.close()
    return rows


def read_xlsx_sections(path: Path) -> dict[str, list[list[str]]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sections: dict[str, list[list[str]]] = {}
    for sheet in workbook.worksheets:
        rows = []
        for row in sheet.iter_rows(values_only=True):
            rows.append([_normalize_cell(cell) for cell in row])
        sections[sheet.title] = rows
    workbook.close()
    return sections


def _build_line_entries(lines: list[str], prefix: str = "Riga") -> list[TextEntry]:
    return [TextEntry(text=line, location=f"{prefix} {index}") for index, line in enumerate(lines, start=1)]


def _paragraph_has_page_break(paragraph) -> bool:
    for run in paragraph.runs:
        for line_break in run._element.findall(".//w:br", run._element.nsmap):
            if line_break.get(qn("w:type")) == "page":
                return True
    return False


def _iter_docx_blocks(document) -> Iterator[Paragraph | Table]:
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


def read_docx_entries(path: Path) -> list[TextEntry]:
    document = Document(path)
    entries: list[TextEntry] = []
    estimated_page = 1
    paragraph_in_page = 0
    table_index = 0
    latest_paragraph_title = ""

    for block in _iter_docx_blocks(document):
        if isinstance(block, Paragraph):
            paragraph = block
        else:
            table = block
            table_index += 1
            table_title = latest_paragraph_title or f"Tabella {table_index}"
            table_row_number = 0
            for row in table.rows:
                row_values = [cell.text.strip() for cell in row.cells]
                if not any(row_values):
                    continue
                table_row_number += 1
                for column_number, cell_value in enumerate(row_values, start=1):
                    entries.append(
                        TextEntry(
                            text=cell_value,
                            location=(
                                f'Pag. {estimated_page} (stimata), tabella "{table_title}", '
                                f"riga {table_row_number}, colonna {column_number}"
                            ),
                        )
                    )
            continue

        if _paragraph_has_page_break(paragraph):
            estimated_page += 1
            paragraph_in_page = 0

        text = paragraph.text.strip()
        if not text:
            continue

        paragraph_in_page += 1
        latest_paragraph_title = text
        entries.append(
            TextEntry(
                text=text,
                location=f"Pag. {estimated_page} (stimata), paragrafo {paragraph_in_page}",
            )
        )

    return entries


def _extract_table_cells(line: str) -> list[str] | None:
    if "|" in line:
        cells = [cell.strip() for cell in line.split("|")]
    elif "\t" in line:
        cells = [cell.strip() for cell in line.split("\t")]
    else:
        cells = [cell.strip() for cell in re.split(r"\s{2,}", line)]

    non_empty_cells = [cell for cell in cells if cell]
    if len(non_empty_cells) < 2:
        return None
    return cells


def read_pdf_entries(path: Path) -> list[TextEntry]:
    reader = PdfReader(str(path))
    entries: list[TextEntry] = []
    table_index = 0
    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        active_table_row = 0
        active_table_title = ""
        previous_plain_line = ""
        for line_number, line in enumerate(text.splitlines(), start=1):
            stripped_line = line.strip()
            if not stripped_line:
                continue
            cells = _extract_table_cells(stripped_line)
            if cells is not None:
                if active_table_row == 0:
                    table_index += 1
                    active_table_title = previous_plain_line or f"Tabella {table_index}"
                active_table_row += 1
                for column_number, cell_value in enumerate(cells, start=1):
                    entries.append(
                        TextEntry(
                            text=cell_value,
                            location=(
                                f'Pag. {page_number}, tabella "{active_table_title}", '
                                f"riga {active_table_row}, colonna {column_number}"
                            ),
                        )
                    )
                continue

            active_table_row = 0
            entries.append(
                TextEntry(
                    text=stripped_line,
                    location=f"Pag. {page_number}, riga {line_number}",
                )
            )
            previous_plain_line = stripped_line
    return entries


def read_doc_lines(path: Path) -> list[str]:
    raise ComparisonError(
        "I file .doc legacy non sono supportati direttamente. Converti il file in .docx oppure .pdf e riprova."
    )


def read_text_lines(path: Path) -> list[str]:
    return [entry.text for entry in read_text_entries(path)]


def read_text_entries(path: Path) -> list[TextEntry]:
    extension = path.suffix.lower()
    if extension == ".txt":
        with path.open("r", encoding="utf-8-sig") as handle:
            lines = [line.rstrip("\n") for line in handle]
        return _build_line_entries(lines)
    if extension == ".csv":
        return _build_line_entries(["\t".join(row) for row in read_csv_rows(path)])
    if extension in {".xlsx", ".xlsm"}:
        sections = read_xlsx_sections(path)
        lines: list[str] = []
        include_sheet_markers = len(sections) > 1
        for sheet_name, rows in sections.items():
            if include_sheet_markers:
                lines.append(f"[{sheet_name}]")
            lines.extend("\t".join(row) for row in rows)
        return _build_line_entries(lines)
    if extension == ".pdf":
        return read_pdf_entries(path)
    if extension == ".docx":
        return read_docx_entries(path)
    if extension == ".doc":
        return read_doc_lines(path)

    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            lines = [line.rstrip("\n") for line in handle]
        return _build_line_entries(lines)
    except UnicodeDecodeError as exc:
        raise ComparisonError(f"Formato non supportato per la lettura testuale: {path.suffix}") from exc


def read_tabular_rows(path: Path) -> list[list[str]]:
    sections = read_tabular_sections(path)
    if not sections:
        return []
    return next(iter(sections.values()))


def read_tabular_sections(path: Path) -> dict[str, list[list[str]]]:
    extension = path.suffix.lower()
    if extension == ".csv":
        return {DEFAULT_TABULAR_SECTION: read_csv_rows(path)}
    if extension in {".xlsx", ".xlsm"}:
        return read_xlsx_sections(path)
    raise ComparisonError(f"Formato tabellare non supportato: {path.suffix}")


def split_header(rows1: list[list[str]], rows2: list[list[str]]) -> tuple[list[str], list[list[str]], list[list[str]]]:
    if rows1 and rows2 and len(rows1[0]) == len(rows2[0]) and rows1[0] == rows2[0]:
        header = [cell or f"Colonna {index + 1}" for index, cell in enumerate(rows1[0])]
        return header, rows1[1:], rows2[1:]

    width = max((len(row) for row in rows1 + rows2), default=0)
    header = [f"Colonna {index + 1}" for index in range(width)]
    return header, rows1, rows2


def pad_row(row: Sequence[str], width: int) -> list[str]:
    return list(row) + [""] * (width - len(row))


def build_record_key(row: Sequence[str], row_number: int, key_indexes: tuple[int, ...] | None) -> str:
    if key_indexes is None:
        return str(row_number)
    try:
        return " | ".join(row[index] for index in key_indexes)
    except IndexError as exc:
        raise ComparisonError("La chiave fa riferimento a colonne non presenti nel file.") from exc


def rows_to_mapping(rows: list[list[str]], key_indexes: tuple[int, ...] | None, width: int) -> dict[str, list[str]]:
    mapping: dict[str, list[str]] = {}
    for offset, raw_row in enumerate(rows, start=1):
        row = pad_row(raw_row, width)
        record_key = build_record_key(row, offset, key_indexes)
        if key_indexes is not None and not record_key.replace("|", "").strip():
            continue
        if record_key in mapping:
            raise ComparisonError(f"Chiave duplicata trovata nel file: {record_key}")
        mapping[record_key] = row
    return mapping


def compare_tabular_files(file1: Path, file2: Path, key_indexes: tuple[int, ...] | None) -> ComparisonResult:
    sections1 = read_tabular_sections(file1)
    sections2 = read_tabular_sections(file2)
    compare_single_section = len(sections1) == 1 and len(sections2) == 1

    if compare_single_section:
        only_rows1 = next(iter(sections1.values()))
        only_rows2 = next(iter(sections2.values()))
        sections_to_compare = [(None, only_rows1, only_rows2)]
    else:
        section_names = sorted(set(sections1) | set(sections2))
        sections_to_compare = [
            (name, sections1.get(name, []), sections2.get(name, [])) for name in section_names
        ]

    differences: list[dict[str, str | int]] = []
    added = removed = changed = 0
    records_file1 = records_file2 = 0

    for section_name, rows1, rows2 in sections_to_compare:
        if rows1 and not rows2:
            header = [cell or f"Colonna {index + 1}" for index, cell in enumerate(rows1[0])]
            data1, data2 = rows1[1:], []
        elif rows2 and not rows1:
            header = [cell or f"Colonna {index + 1}" for index, cell in enumerate(rows2[0])]
            data1, data2 = [], rows2[1:]
        else:
            header, data1, data2 = split_header(rows1, rows2)

        width = len(header)
        mapping1 = rows_to_mapping(data1, key_indexes, width)
        mapping2 = rows_to_mapping(data2, key_indexes, width)
        records_file1 += len(data1)
        records_file2 += len(data2)

        for record_key in sorted(set(mapping1) | set(mapping2)):
            row1 = mapping1.get(record_key)
            row2 = mapping2.get(record_key)

            if row1 is None and row2 is not None:
                added += 1
                for column_name, value2 in zip(header, row2):
                    difference: dict[str, str | int] = {
                        "status": "ADDED",
                        "record_key": record_key,
                        "column": column_name,
                        "file1": "",
                        "file2": value2,
                    }
                    if section_name is not None:
                        difference["sheet"] = section_name
                    differences.append(difference)
                continue

            if row2 is None and row1 is not None:
                removed += 1
                for column_name, value1 in zip(header, row1):
                    difference = {
                        "status": "REMOVED",
                        "record_key": record_key,
                        "column": column_name,
                        "file1": value1,
                        "file2": "",
                    }
                    if section_name is not None:
                        difference["sheet"] = section_name
                    differences.append(difference)
                continue

            assert row1 is not None and row2 is not None
            row_changed = False
            for column_name, value1, value2 in zip(header, row1, row2):
                if value1 != value2:
                    row_changed = True
                    difference = {
                        "status": "CHANGED",
                        "record_key": record_key,
                        "column": column_name,
                        "file1": value1,
                        "file2": value2,
                    }
                    if section_name is not None:
                        difference["sheet"] = section_name
                    differences.append(difference)
            if row_changed:
                changed += 1

    return ComparisonResult(
        mode="tabular",
        differences=differences,
        summary={
            "file1": str(file1),
            "file2": str(file2),
            "key": "auto-riga" if key_indexes is None else "+".join(str(index + 1) for index in key_indexes),
            "records_file1": records_file1,
            "records_file2": records_file2,
            "added": added,
            "removed": removed,
            "changed": changed,
        },
    )


def compare_text_files(file1: Path, file2: Path) -> ComparisonResult:
    entries1 = read_text_entries(file1)
    entries2 = read_text_entries(file2)
    lines1 = [entry.text for entry in entries1]
    lines2 = [entry.text for entry in entries2]
    matcher = SequenceMatcher(a=lines1, b=lines2)
    differences: list[dict[str, str | int]] = []
    added = removed = changed = 0

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "delete":
            for offset, line in enumerate(lines1[i1:i2], start=i1 + 1):
                removed += 1
                differences.append(
                    {
                        "status": "REMOVED",
                        "line": offset,
                        "location_file1": entries1[offset - 1].location,
                        "context_file1": build_entry_context(entries1, offset - 1),
                        "location_file2": "",
                        "context_file2": "",
                        "file1": line,
                        "file2": "",
                    }
                )
            continue
        if tag == "insert":
            for offset, line in enumerate(lines2[j1:j2], start=j1 + 1):
                added += 1
                differences.append(
                    {
                        "status": "ADDED",
                        "line": offset,
                        "location_file1": "",
                        "context_file1": "",
                        "location_file2": entries2[offset - 1].location,
                        "context_file2": build_entry_context(entries2, offset - 1),
                        "file1": "",
                        "file2": line,
                    }
                )
            continue

        left = lines1[i1:i2]
        right = lines2[j1:j2]
        max_len = max(len(left), len(right))
        for index in range(max_len):
            changed += 1
            left_index = i1 + index
            right_index = j1 + index
            differences.append(
                {
                    "status": "CHANGED",
                    "line": max(i1, j1) + index + 1,
                    "location_file1": entries1[left_index].location if left_index < len(entries1) else "",
                    "context_file1": build_entry_context(entries1, left_index) if left_index < len(entries1) else "",
                    "location_file2": entries2[right_index].location if right_index < len(entries2) else "",
                    "context_file2": build_entry_context(entries2, right_index) if right_index < len(entries2) else "",
                    "file1": left[index] if index < len(left) else "",
                    "file2": right[index] if index < len(right) else "",
                }
            )

    return ComparisonResult(
        mode="text",
        differences=differences,
        summary={
            "file1": str(file1),
            "file2": str(file2),
            "lines_file1": len(lines1),
            "lines_file2": len(lines2),
            "added": added,
            "removed": removed,
            "changed": changed,
        },
    )


def auto_compare(file1: Path, file2: Path, key_spec: str | None = None) -> ComparisonResult:
    key_indexes = parse_key_spec(key_spec)
    both_tabular = file1.suffix.lower() in TABULAR_EXTENSIONS and file2.suffix.lower() in TABULAR_EXTENSIONS
    if both_tabular:
        return compare_tabular_files(file1, file2, key_indexes)
    if key_indexes is not None:
        raise ComparisonError("Il parametro --key è disponibile solo per confronti tabellari CSV/XLSX.")
    return compare_text_files(file1, file2)


def build_details_table(result: ComparisonResult) -> tuple[list[str], list[list[str | int]]]:
    if result.mode == "tabular":
        include_sheet = any("sheet" in difference for difference in result.differences)
        headers = ["Status", "Record Key", "Column", "File 1", "File 2"]
        if include_sheet:
            headers.insert(1, "Sheet")

        rows = []
        for difference in result.differences:
            row: list[str | int] = [
                difference["status"],
                difference["record_key"],
                difference["column"],
                difference["file1"],
                difference["file2"],
            ]
            if include_sheet:
                row.insert(1, difference.get("sheet", ""))
            rows.append(row)
        return headers, rows

    headers = [
        "Status",
        "Line",
        "Posizione File 1",
        "Contesto File 1",
        "Posizione File 2",
        "Contesto File 2",
        "File 1",
        "File 2",
    ]
    rows = [
        [
            difference["status"],
            difference["line"],
            difference.get("location_file1", ""),
            difference.get("context_file1", ""),
            difference.get("location_file2", ""),
            difference.get("context_file2", ""),
            difference["file1"],
            difference["file2"],
        ]
        for difference in result.differences
    ]
    return headers, rows


def write_csv_report(result: ComparisonResult, output_path: Path) -> Path:
    csv_path = output_path.with_suffix(".csv")
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    headers, rows = build_details_table(result)

    with csv_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Campo", "Valore"])
        for field, value in result.summary.items():
            writer.writerow([field, value])
        writer.writerow([])
        writer.writerow(headers)
        writer.writerows(rows)

    return csv_path


def write_excel_report(result: ComparisonResult, output_path: Path) -> Path:
    headers, detail_rows = build_details_table(result)
    if len(detail_rows) + 1 > MAX_EXCEL_ROWS:
        return write_csv_report(result, output_path)

    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Summary"
    summary_sheet.append(["Campo", "Valore"])
    for field, value in result.summary.items():
        summary_sheet.append([field, value])
    summary_sheet.freeze_panes = "A2"

    details_sheet = workbook.create_sheet("Differences")
    details_sheet.append(headers)
    for row in detail_rows:
        details_sheet.append(row)

    details_sheet.freeze_panes = "A2"
    details_sheet.auto_filter.ref = details_sheet.dimensions
    for row in details_sheet.iter_rows(min_row=2):
        status = row[0].value
        fill = STATUS_FILLS.get(status)
        if fill:
            for cell in row:
                cell.fill = fill

    for sheet in (summary_sheet, details_sheet):
        for column_cells in sheet.columns:
            length = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 60)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Confronta due file e genera un report Excel con le differenze."
    )
    parser.add_argument("file1", type=Path, help="Primo file da confrontare")
    parser.add_argument("file2", type=Path, help="Secondo file da confrontare")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("differenze.xlsx"),
        help="Percorso del file Excel di output",
    )
    parser.add_argument(
        "-k",
        "--key",
        dest="key_spec",
        help="Chiave record per confronti tabellari, nel formato 1+5",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.file1.exists() or not args.file2.exists():
        parser.error("Entrambi i file di input devono esistere.")

    try:
        result = auto_compare(args.file1, args.file2, args.key_spec)
        report_path = write_excel_report(result, args.output)
    except ComparisonError as exc:
        parser.error(str(exc))

    print(f"Report generato: {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
