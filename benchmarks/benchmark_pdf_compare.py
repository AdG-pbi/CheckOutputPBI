from __future__ import annotations

import argparse
import statistics
import tempfile
from pathlib import Path
from time import perf_counter

import pymupdf as fitz

from compare_files import auto_compare, write_highlight_pdf_for_file2


def _build_text_pdf(path: Path, *, pages: int, lines_per_page: int, changed_suffix: str = "") -> None:
    document = fitz.open()
    for page_index in range(pages):
        page = document.new_page()
        for line_index in range(lines_per_page):
            page.insert_text(
                (36, 36 + line_index * 9),
                f"Pag {page_index + 1} Riga {line_index + 1}{changed_suffix if line_index % 25 == 0 else ''}",
                fontsize=8,
            )
    document.save(path)
    document.close()


def _build_visual_only_pdf(path: Path, *, pages: int, lines_per_page: int, add_markers: bool) -> None:
    document = fitz.open()
    for page_index in range(pages):
        page = document.new_page()
        for line_index in range(lines_per_page):
            page.insert_text((36, 36 + line_index * 9), f"Visual page {page_index + 1} line {line_index + 1}", fontsize=8)
        if add_markers:
            page.draw_rect((420, 36, 500, 120), color=(0, 0, 0), fill=(0, 0, 0))
    document.save(path)
    document.close()


def _run_case(name: str, left: Path, right: Path, output: Path, iterations: int) -> None:
    compare_timings: list[float] = []
    highlight_timings: list[float] = []
    for _ in range(iterations):
        start = perf_counter()
        _ = auto_compare(left, right)
        compare_timings.append(perf_counter() - start)

        start = perf_counter()
        write_highlight_pdf_for_file2(left, right, output)
        highlight_timings.append(perf_counter() - start)

    compare_avg = statistics.mean(compare_timings)
    highlight_avg = statistics.mean(highlight_timings)
    total_avg = compare_avg + highlight_avg
    print(
        f"{name:18} compare={compare_avg:.3f}s highlight={highlight_avg:.3f}s total={total_avg:.3f}s "
        f"(iterazioni={iterations})"
    )


def run_benchmarks(iterations: int) -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)

        small_left = tmp_path / "small_left.pdf"
        small_right = tmp_path / "small_right.pdf"
        _build_text_pdf(small_left, pages=1, lines_per_page=40)
        _build_text_pdf(small_right, pages=1, lines_per_page=40, changed_suffix=" changed")
        _run_case("small", small_left, small_right, tmp_path / "small_highlight.pdf", iterations)

        medium_left = tmp_path / "medium_left.pdf"
        medium_right = tmp_path / "medium_right.pdf"
        _build_text_pdf(medium_left, pages=5, lines_per_page=90)
        _build_text_pdf(medium_right, pages=5, lines_per_page=90, changed_suffix=" changed")
        _run_case("medium", medium_left, medium_right, tmp_path / "medium_highlight.pdf", iterations)

        large_left = tmp_path / "large_left.pdf"
        large_right = tmp_path / "large_right.pdf"
        _build_text_pdf(large_left, pages=20, lines_per_page=120)
        _build_text_pdf(large_right, pages=20, lines_per_page=120, changed_suffix=" changed")
        _run_case("large", large_left, large_right, tmp_path / "large_highlight.pdf", iterations)

        visual_left = tmp_path / "visual_left.pdf"
        visual_right = tmp_path / "visual_right.pdf"
        _build_visual_only_pdf(visual_left, pages=5, lines_per_page=80, add_markers=False)
        _build_visual_only_pdf(visual_right, pages=5, lines_per_page=80, add_markers=True)
        _run_case("visual-only", visual_left, visual_right, tmp_path / "visual_highlight.pdf", iterations)


def main() -> int:
    parser = argparse.ArgumentParser(description="Benchmark PDF comparison and highlight generation.")
    parser.add_argument(
        "--iterations",
        type=int,
        default=3,
        help="Numero di iterazioni per scenario (default: 3).",
    )
    args = parser.parse_args()
    run_benchmarks(args.iterations)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
