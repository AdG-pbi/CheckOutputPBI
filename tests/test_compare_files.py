import csv
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.enum.text import WD_BREAK
from openpyxl import Workbook, load_workbook

from compare_files import (
    ComparisonResult,
    auto_compare,
    build_file2_highlight_lines,
    parse_key_spec,
    read_xlsx_rows,
    write_excel_report,
    write_highlight_pdf_for_file2,
)


class CompareFilesTests(unittest.TestCase):
    def test_parse_key_spec(self):
        self.assertEqual(parse_key_spec("1+5"), (0, 4))
        self.assertEqual(parse_key_spec(None), None)

    def test_compare_csv_with_key_and_excel_report(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.csv"
            right = tmp_path / "right.csv"
            report = tmp_path / "report.xlsx"

            with left.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerows(
                    [
                        ["id", "name", "city"],
                        ["1", "Alice", "Rome"],
                        ["2", "Bob", "Milan"],
                    ]
                )

            with right.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerows(
                    [
                        ["id", "name", "city"],
                        ["1", "Alice", "Turin"],
                        ["3", "Carla", "Naples"],
                    ]
                )

            result = auto_compare(left, right, "1")
            write_excel_report(result, report)

            self.assertEqual(result.summary["changed"], 1)
            self.assertEqual(result.summary["added"], 1)
            self.assertEqual(result.summary["removed"], 1)

            workbook = load_workbook(report)
            sheet = workbook["Differences"]
            rows = list(sheet.iter_rows(values_only=True))
            workbook.close()

            self.assertIn(("CHANGED", "1", "city", "Rome", "Turin"), rows)
            self.assertIn(("REMOVED", "2", "id", "2", None), rows)
            self.assertIn(("ADDED", "3", "name", None, "Carla"), rows)

    def test_compare_text_files(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.txt"
            right = tmp_path / "right.txt"
            report = tmp_path / "report.xlsx"

            left.write_text("alpha\nbeta\ngamma\n", encoding="utf-8")
            right.write_text("alpha\nbeta 2\ndelta\n", encoding="utf-8")

            result = auto_compare(left, right)
            write_excel_report(result, report)

            self.assertEqual(result.mode, "text")
            self.assertEqual(result.summary["changed"], 2)
            self.assertEqual(result.differences[0]["status"], "CHANGED")
            self.assertEqual(result.differences[0]["location_file1"], "Riga 2")
            self.assertEqual(result.differences[0]["location_file2"], "Riga 2")
            self.assertEqual(result.differences[0]["context_file1"], "Dopo: alpha | Prima di: gamma")
            self.assertEqual(result.differences[0]["context_file2"], "Dopo: alpha | Prima di: delta")

            workbook = load_workbook(report)
            sheet = workbook["Differences"]
            rows = list(sheet.iter_rows(values_only=True))
            workbook.close()

            self.assertEqual(
                rows[0],
                (
                    "Status",
                    "Line",
                    "Posizione File 1",
                    "Contesto File 1",
                    "Posizione File 2",
                    "Contesto File 2",
                    "File 1",
                    "File 2",
                ),
            )
            self.assertIn(
                (
                    "CHANGED",
                    2,
                    "Riga 2",
                    "Dopo: alpha | Prima di: gamma",
                    "Riga 2",
                    "Dopo: alpha | Prima di: delta",
                    "beta",
                    "beta 2",
                ),
                rows,
            )

    def test_build_file2_highlight_lines_marks_changed_characters_and_numbers(self):
        highlighted = build_file2_highlight_lines(["Numero: 123"], ["Numero: 129"])

        self.assertEqual(len(highlighted), 1)
        self.assertEqual(highlighted[0], [("Numero: 12", False), ("9", True)])

    def test_write_highlight_pdf_for_docx_file2_creates_pdf_output(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.docx"
            right = tmp_path / "right.docx"
            output_pdf = tmp_path / "highlighted.pdf"

            doc_left = Document()
            doc_left.add_paragraph("Valore A")
            doc_left.save(left)

            doc_right = Document()
            doc_right.add_paragraph("Valore B")
            doc_right.save(right)

            actual_path = write_highlight_pdf_for_file2(left, right, output_pdf)

            self.assertEqual(actual_path, output_pdf)
            self.assertTrue(output_pdf.exists())
            self.assertGreater(output_pdf.stat().st_size, 0)

    def test_write_highlight_pdf_for_pdf_file2_keeps_original_pdf_and_adds_only_changed_line_highlights(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.pdf"
            right = tmp_path / "right.pdf"
            output_pdf = tmp_path / "highlighted.pdf"
            left.write_bytes(b"%PDF-1.7")
            right.write_bytes(b"%PDF-1.7")

            class FakePage:
                def __init__(self, text):
                    self._text = text
                    self.highlights = []

                def get_text(self, _mode):
                    return self._text

                def search_for(self, text):
                    if text in self._text:
                        return [f"rect::{text}"]
                    return []

                def add_highlight_annot(self, rect):
                    self.highlights.append(rect)
                    return rect

            class FakeDocument:
                def __init__(self, pages, source_path):
                    self.pages = pages
                    self.source_path = source_path

                def __enter__(self):
                    return self

                def __exit__(self, exc_type, exc, tb):
                    return False

                def __iter__(self):
                    return iter(self.pages)

                def __getitem__(self, index):
                    return self.pages[index]

                def save(self, destination):
                    Path(destination).write_bytes(self.source_path.read_bytes())

            left_document = FakeDocument([FakePage("Riga invariata\nRiga A")], left)
            right_page = FakePage("Riga invariata\nRiga B")
            right_document = FakeDocument([right_page], right)

            def fake_open(path):
                path = Path(path)
                if path == left:
                    return left_document
                if path == right:
                    return right_document
                self.fail(f"Unexpected path passed to fitz.open: {path}")

            with patch("compare_files.fitz") as mock_fitz:
                mock_fitz.open.side_effect = fake_open
                actual_path = write_highlight_pdf_for_file2(left, right, output_pdf)

            self.assertEqual(actual_path, output_pdf)
            self.assertTrue(output_pdf.exists())
            self.assertEqual(output_pdf.read_bytes(), right.read_bytes())
            self.assertEqual(right_page.highlights, ["rect::Riga B"])

    def test_compare_pdf_files_exposes_page_and_line_locations(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.pdf"
            right = tmp_path / "right.pdf"
            left.write_bytes(b"%PDF-1.7")
            right.write_bytes(b"%PDF-1.7")

            class FakePage:
                def __init__(self, text):
                    self._text = text

                def extract_text(self, *args, **kwargs):
                    return self._text

            class FakeReader:
                def __init__(self, pages):
                    self.pages = pages

            with patch(
                "compare_files.PdfReader",
                side_effect=[
                    FakeReader([FakePage("Titolo\nValore A"), FakePage("Conclusione")]),
                    FakeReader([FakePage("Titolo\nValore B"), FakePage("Conclusione")]),
                ],
            ):
                result = auto_compare(left, right)

            changed_rows = [row for row in result.differences if row["status"] == "CHANGED"]
            self.assertEqual(len(changed_rows), 1)
            self.assertEqual(changed_rows[0]["location_file1"], "Pag. 1, riga 2")
            self.assertEqual(changed_rows[0]["location_file2"], "Pag. 1, riga 2")
            self.assertEqual(changed_rows[0]["context_file1"], "Dopo: Titolo | Prima di: Conclusione")
            self.assertEqual(changed_rows[0]["context_file2"], "Dopo: Titolo | Prima di: Conclusione")

    def test_compare_docx_files_exposes_estimated_page_locations(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.docx"
            right = tmp_path / "right.docx"

            doc_left = Document()
            doc_left.add_paragraph("Introduzione")
            break_paragraph_left = doc_left.add_paragraph()
            break_paragraph_left.add_run().add_break(WD_BREAK.PAGE)
            doc_left.add_paragraph("Valore A")
            doc_left.save(left)

            doc_right = Document()
            doc_right.add_paragraph("Introduzione")
            break_paragraph_right = doc_right.add_paragraph()
            break_paragraph_right.add_run().add_break(WD_BREAK.PAGE)
            doc_right.add_paragraph("Valore B")
            doc_right.save(right)

            result = auto_compare(left, right)

            changed_rows = [row for row in result.differences if row["status"] == "CHANGED"]
            self.assertEqual(len(changed_rows), 1)
            self.assertEqual(changed_rows[0]["location_file1"], "Pag. 2 (stimata), paragrafo 1")
            self.assertEqual(changed_rows[0]["location_file2"], "Pag. 2 (stimata), paragrafo 1")
            self.assertEqual(changed_rows[0]["context_file1"], "Dopo: Introduzione")
            self.assertEqual(changed_rows[0]["context_file2"], "Dopo: Introduzione")

    def test_compare_docx_table_files_expose_table_title_row_and_column(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.docx"
            right = tmp_path / "right.docx"

            doc_left = Document()
            doc_left.add_paragraph("Tabella Clienti")
            table_left = doc_left.add_table(rows=3, cols=2)
            table_left.cell(0, 0).text = "ID"
            table_left.cell(0, 1).text = "Nome"
            table_left.cell(1, 0).text = "1"
            table_left.cell(1, 1).text = "Alice"
            table_left.cell(2, 0).text = "2"
            table_left.cell(2, 1).text = "Bob"
            doc_left.save(left)

            doc_right = Document()
            doc_right.add_paragraph("Tabella Clienti")
            table_right = doc_right.add_table(rows=3, cols=2)
            table_right.cell(0, 0).text = "ID"
            table_right.cell(0, 1).text = "Nome"
            table_right.cell(1, 0).text = "1"
            table_right.cell(1, 1).text = "Alice"
            table_right.cell(2, 0).text = "2"
            table_right.cell(2, 1).text = "Bobby"
            doc_right.save(right)

            result = auto_compare(left, right)

            changed_rows = [row for row in result.differences if row["status"] == "CHANGED"]
            self.assertEqual(len(changed_rows), 1)
            self.assertEqual(
                changed_rows[0]["location_file1"],
                'Pag. 1 (stimata), tabella "Tabella Clienti", riga 3, colonna 2',
            )
            self.assertEqual(
                changed_rows[0]["location_file2"],
                'Pag. 1 (stimata), tabella "Tabella Clienti", riga 3, colonna 2',
            )

    def test_compare_pdf_table_files_expose_table_title_row_and_column(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.pdf"
            right = tmp_path / "right.pdf"
            left.write_bytes(b"%PDF-1.7")
            right.write_bytes(b"%PDF-1.7")

            class FakePage:
                def __init__(self, text):
                    self._text = text

                def extract_text(self, *args, **kwargs):
                    return self._text

            class FakeReader:
                def __init__(self, pages):
                    self.pages = pages

            with patch(
                "compare_files.PdfReader",
                side_effect=[
                    FakeReader([FakePage("Tabella Clienti\nID | Nome\n1 | Alice\n2 | Bob")]),
                    FakeReader([FakePage("Tabella Clienti\nID | Nome\n1 | Alice\n2 | Bobby")]),
                ],
            ):
                result = auto_compare(left, right)

            changed_rows = [row for row in result.differences if row["status"] == "CHANGED"]
            self.assertEqual(len(changed_rows), 1)
            self.assertEqual(
                changed_rows[0]["location_file1"],
                'Pag. 1, tabella "Tabella Clienti", riga 3, colonna 2',
            )
            self.assertEqual(
                changed_rows[0]["location_file2"],
                'Pag. 1, tabella "Tabella Clienti", riga 3, colonna 2',
            )

    def test_compare_pdf_table_files_use_layout_extraction_when_plain_text_is_flattened(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.pdf"
            right = tmp_path / "right.pdf"
            left.write_bytes(b"%PDF-1.7")
            right.write_bytes(b"%PDF-1.7")

            class FakePage:
                def __init__(self, plain_text, layout_text):
                    self._plain_text = plain_text
                    self._layout_text = layout_text

                def extract_text(self, *args, **kwargs):
                    if kwargs.get("extraction_mode") == "layout":
                        return self._layout_text
                    return self._plain_text

            class FakeReader:
                def __init__(self, pages):
                    self.pages = pages

            with patch(
                "compare_files.PdfReader",
                side_effect=[
                    FakeReader(
                        [
                            FakePage(
                                "Tabella Clienti\nID Nome\n1 Alice\n2 Bob",
                                "Tabella Clienti\nID      Nome\n1       Alice\n2       Bob",
                            )
                        ]
                    ),
                    FakeReader(
                        [
                            FakePage(
                                "Tabella Clienti\nID Nome\n1 Alice\n2 Bobby",
                                "Tabella Clienti\nID      Nome\n1       Alice\n2       Bobby",
                            )
                        ]
                    ),
                ],
            ):
                result = auto_compare(left, right)

            changed_rows = [row for row in result.differences if row["status"] == "CHANGED"]
            self.assertEqual(len(changed_rows), 1)
            self.assertEqual(
                changed_rows[0]["location_file1"],
                'Pag. 1, tabella "Tabella Clienti", riga 3, colonna 2',
            )
            self.assertEqual(
                changed_rows[0]["location_file2"],
                'Pag. 1, tabella "Tabella Clienti", riga 3, colonna 2',
            )

    def test_integer_float_equivalence_in_xlsx(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.xlsx"
            right = tmp_path / "right.xlsx"

            wb1 = Workbook()
            ws1 = wb1.active
            ws1.append(["id", "value"])
            ws1.append([1, 0])
            wb1.save(left)

            wb2 = Workbook()
            ws2 = wb2.active
            ws2.append(["id", "value"])
            ws2.append([1, 0.0])
            wb2.save(right)

            result = auto_compare(left, right, "1")

            self.assertEqual(result.summary["changed"], 0)
            self.assertEqual(result.summary["added"], 0)
            self.assertEqual(result.summary["removed"], 0)

    def test_write_excel_report_falls_back_to_csv_over_limit(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.csv"
            right = tmp_path / "right.csv"
            report = tmp_path / "report.xlsx"

            with left.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerows(
                    [
                        ["id", "name", "city"],
                        ["1", "Alice", "Rome"],
                        ["2", "Bob", "Milan"],
                    ]
                )

            with right.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerows(
                    [
                        ["id", "name", "city"],
                        ["1", "Alice", "Turin"],
                        ["3", "Carla", "Naples"],
                    ]
                )

            result = auto_compare(left, right, "1")

            with patch("compare_files.MAX_EXCEL_ROWS", 5):
                actual_report = write_excel_report(result, report)

            self.assertEqual(actual_report, report.with_suffix(".csv"))
            self.assertFalse(report.exists())
            self.assertTrue(actual_report.exists())

            with actual_report.open("r", encoding="utf-8", newline="") as handle:
                rows = list(csv.reader(handle))

            self.assertEqual(rows[0], ["Campo", "Valore"])
            self.assertIn(["changed", "1"], rows)
            self.assertIn(["Status", "Record Key", "Column", "File 1", "File 2"], rows)
            self.assertIn(["CHANGED", "1", "city", "Rome", "Turin"], rows)

    def test_compare_multi_sheet_xlsx_with_sheet_context(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.xlsx"
            right = tmp_path / "right.xlsx"
            report = tmp_path / "report.xlsx"

            wb1 = Workbook()
            ws1 = wb1.active
            ws1.title = "Clienti"
            ws1.append(["id", "name"])
            ws1.append([1, "Alice"])

            ws2 = wb1.create_sheet("Ordini")
            ws2.append(["id", "total"])
            ws2.append([100, 50])
            wb1.save(left)

            wb2 = Workbook()
            ws1_b = wb2.active
            ws1_b.title = "Clienti"
            ws1_b.append(["id", "name"])
            ws1_b.append([1, "Alicia"])

            ws2_b = wb2.create_sheet("Ordini")
            ws2_b.append(["id", "total"])
            ws2_b.append([100, 50])
            ws2_b.append([200, 10])
            wb2.save(right)

            result = auto_compare(left, right, "1")
            write_excel_report(result, report)

            self.assertEqual(result.summary["changed"], 1)
            self.assertEqual(result.summary["added"], 1)
            self.assertEqual(result.summary["removed"], 0)

            workbook = load_workbook(report)
            sheet = workbook["Differences"]
            rows = list(sheet.iter_rows(values_only=True))
            clienti_rows = list(workbook["Clienti"].iter_rows(values_only=True))
            ordini_rows = list(workbook["Ordini"].iter_rows(values_only=True))
            workbook.close()

            self.assertEqual(rows[0], ("Status", "Sheet", "Record Key", "Column", "File 1", "File 2"))
            self.assertIn(("CHANGED", "Clienti", "1", "name", "Alice", "Alicia"), rows)
            self.assertIn(("ADDED", "Ordini", "200", "total", None, "10"), rows)
            self.assertEqual(clienti_rows[0], ("Status", "Record Key", "Column", "File 1", "File 2"))
            self.assertIn(("CHANGED", "1", "name", "Alice", "Alicia"), clienti_rows)
            self.assertIn(("ADDED", "200", "total", None, "10"), ordini_rows)

    def test_multi_sheet_report_includes_tabs_without_differences(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.xlsx"
            right = tmp_path / "right.xlsx"
            report = tmp_path / "report.xlsx"

            wb1 = Workbook()
            ws1 = wb1.active
            ws1.title = "Clienti"
            ws1.append(["id", "name"])
            ws1.append([1, "Alice"])
            ws2 = wb1.create_sheet("Ordini")
            ws2.append(["id", "total"])
            ws2.append([100, 50])
            wb1.save(left)

            wb2 = Workbook()
            ws1_b = wb2.active
            ws1_b.title = "Clienti"
            ws1_b.append(["id", "name"])
            ws1_b.append([1, "Alicia"])
            ws2_b = wb2.create_sheet("Ordini")
            ws2_b.append(["id", "total"])
            ws2_b.append([100, 50])
            wb2.save(right)

            result = auto_compare(left, right, "1")
            write_excel_report(result, report)

            workbook = load_workbook(report)
            clienti_rows = list(workbook["Clienti"].iter_rows(values_only=True))
            ordini_rows = list(workbook["Ordini"].iter_rows(values_only=True))
            workbook.close()

            self.assertEqual(clienti_rows[0], ("Status", "Record Key", "Column", "File 1", "File 2"))
            self.assertIn(("CHANGED", "1", "name", "Alice", "Alicia"), clienti_rows)
            self.assertEqual(ordini_rows, [("Status", "Record Key", "Column", "File 1", "File 2")])

    def test_multi_sheet_report_maps_rows_even_with_sheet_whitespace(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            report = Path(tmp_dir) / "report.xlsx"
            result = ComparisonResult(
                mode="tabular",
                differences=[
                    {
                        "status": "CHANGED",
                        "sheet": " Clienti ",
                        "record_key": "1",
                        "column": "name",
                        "file1": "Alice",
                        "file2": "Alicia",
                    }
                ],
                summary={"changed": 1},
                sections=["Clienti"],
            )

            write_excel_report(result, report)

            workbook = load_workbook(report)
            clienti_rows = list(workbook["Clienti"].iter_rows(values_only=True))
            workbook.close()

            self.assertIn(("CHANGED", "1", "name", "Alice", "Alicia"), clienti_rows)

    def test_multi_sheet_report_maps_rows_when_section_name_has_whitespace(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            report = Path(tmp_dir) / "report.xlsx"
            result = ComparisonResult(
                mode="tabular",
                differences=[
                    {
                        "status": "CHANGED",
                        "sheet": " Clienti ",
                        "record_key": "1",
                        "column": "name",
                        "file1": "Alice",
                        "file2": "Alicia",
                    }
                ],
                summary={"changed": 1},
                sections=[" Clienti "],
            )

            write_excel_report(result, report)

            workbook = load_workbook(report)
            clienti_rows = list(workbook["Clienti"].iter_rows(values_only=True))
            workbook.close()

            self.assertIn(("CHANGED", "1", "name", "Alice", "Alicia"), clienti_rows)

    def test_compare_multi_sheet_xlsx_matches_sheet_names_with_outer_whitespace(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            left = tmp_path / "left.xlsx"
            right = tmp_path / "right.xlsx"
            report = tmp_path / "report.xlsx"

            wb1 = Workbook()
            ws1 = wb1.active
            ws1.title = "Clienti"
            ws1.append(["id", "name"])
            ws1.append([1, "Alice"])
            ws2 = wb1.create_sheet("Ordini")
            ws2.append(["id", "total"])
            ws2.append([100, 50])
            wb1.save(left)

            wb2 = Workbook()
            ws1_b = wb2.active
            ws1_b.title = " Clienti "
            ws1_b.append(["id", "name"])
            ws1_b.append([1, "Alicia"])
            ws2_b = wb2.create_sheet("Ordini")
            ws2_b.append(["id", "total"])
            ws2_b.append([100, 50])
            wb2.save(right)

            result = auto_compare(left, right, "1")
            write_excel_report(result, report)

            self.assertEqual(result.summary["changed"], 1)
            self.assertEqual(result.summary["added"], 0)
            self.assertEqual(result.summary["removed"], 0)

            workbook = load_workbook(report)
            rows = list(workbook["Differences"].iter_rows(values_only=True))
            clienti_rows = list(workbook["Clienti"].iter_rows(values_only=True))
            workbook.close()

            self.assertIn(("CHANGED", "Clienti", "1", "name", "Alice", "Alicia"), rows)
            self.assertNotIn(("REMOVED", "Clienti", "1", "name", "Alice", None), rows)
            self.assertIn(("CHANGED", "1", "name", "Alice", "Alicia"), clienti_rows)

    def test_multi_sheet_report_keeps_whitespace_distinct_sections_separate(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            report = Path(tmp_dir) / "report.xlsx"
            result = ComparisonResult(
                mode="tabular",
                differences=[
                    {
                        "status": "CHANGED",
                        "sheet": "Clienti",
                        "record_key": "1",
                        "column": "name",
                        "file1": "Alice",
                        "file2": "Alicia",
                    },
                    {
                        "status": "CHANGED",
                        "sheet": " Clienti ",
                        "record_key": "2",
                        "column": "name",
                        "file1": "Bob",
                        "file2": "Bobby",
                    },
                ],
                summary={"changed": 2},
                sections=["Clienti", " Clienti "],
            )

            write_excel_report(result, report)

            workbook = load_workbook(report)
            clienti_rows = list(workbook["Clienti"].iter_rows(values_only=True))
            clienti_rows_2 = list(workbook["Clienti (2)"].iter_rows(values_only=True))
            workbook.close()

            self.assertIn(("CHANGED", "1", "name", "Alice", "Alicia"), clienti_rows)
            self.assertNotIn(("CHANGED", "2", "name", "Bob", "Bobby"), clienti_rows)
            self.assertIn(("CHANGED", "2", "name", "Bob", "Bobby"), clienti_rows_2)
            self.assertNotIn(("CHANGED", "1", "name", "Alice", "Alicia"), clienti_rows_2)

    def test_read_xlsx_rows_includes_all_sheets(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            workbook_path = tmp_path / "multi.xlsx"

            workbook = Workbook()
            first_sheet = workbook.active
            first_sheet.title = "Clienti"
            first_sheet.append(["id", "name"])
            first_sheet.append([1, "Alice"])

            second_sheet = workbook.create_sheet("Ordini")
            second_sheet.append(["id", "total"])
            second_sheet.append([100, 50])
            workbook.save(workbook_path)

            rows = read_xlsx_rows(workbook_path)

            self.assertEqual(
                rows,
                [
                    ["[Clienti]"],
                    ["id", "name"],
                    ["1", "Alice"],
                    ["[Ordini]"],
                    ["id", "total"],
                    ["100", "50"],
                ],
            )

    def test_read_xlsx_rows_skips_empty_rows_and_columns_but_keeps_scanning(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            workbook_path = tmp_path / "sparse.xlsx"

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Dati"
            sheet["A1"] = "id"
            sheet["D1"] = "name"
            sheet["A2"] = 1
            sheet["D2"] = "Alice"
            sheet["A4"] = 2
            sheet["D4"] = "Bob"
            workbook.save(workbook_path)

            rows = read_xlsx_rows(workbook_path)

            self.assertEqual(
                rows,
                [
                    ["id", "name"],
                    ["1", "Alice"],
                    ["2", "Bob"],
                ],
            )


if __name__ == "__main__":
    unittest.main()
