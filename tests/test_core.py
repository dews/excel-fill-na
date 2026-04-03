from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_fillna.core import fill_empty_cells, process_workbook


def test_fill_empty_cells_only_in_selected_range() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet["A2"] = "keep"

    result = fill_empty_cells(worksheet, target_range="A1:A2")

    assert worksheet["A1"].value == "NA"
    assert worksheet["A2"].value == "keep"
    assert worksheet["B1"].value is None
    assert result.filled_cells == 1
    assert result.merged_ranges == ()


def test_fill_empty_cells_respects_excluded_ranges() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet["B2"] = "keep"

    result = fill_empty_cells(
        worksheet,
        target_range="A1:B2",
        excluded_ranges=["A2", "B1:B2"],
        fill_value="MISSING",
    )

    assert worksheet["A1"].value == "MISSING"
    assert worksheet["A2"].value is None
    assert worksheet["B1"].value is None
    assert worksheet["B2"].value == "keep"
    assert result.filled_cells == 1


def test_merge_empty_runs_creates_vertical_merge() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet["A4"] = "stop"

    result = fill_empty_cells(
        worksheet,
        target_range="A1:A4",
        merge_empty_runs=True,
    )

    merged_ranges = {str(cell_range) for cell_range in worksheet.merged_cells.ranges}

    assert "A1:A3" in merged_ranges
    assert worksheet["A1"].value == "NA"
    assert worksheet["A4"].value == "stop"
    assert result.filled_cells == 3
    assert result.merged_ranges == ("A1:A3",)


def test_existing_merged_cells_are_supported() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.merge_cells("B1:B3")

    result = fill_empty_cells(
        worksheet,
        target_range="B1:B3",
        fill_value="EMPTY",
    )

    assert worksheet["B1"].value == "EMPTY"
    assert result.filled_cells == 1
    assert result.merged_ranges == ()


def test_process_workbook_saves_default_output_file(tmp_path: Path) -> None:
    source = tmp_path / "input.xlsx"
    workbook = Workbook()
    workbook.active["C3"] = None
    workbook.save(source)

    result = process_workbook(
        source,
        target_range="C3",
        fill_value="EMPTY",
    )

    assert result.output_path == tmp_path / "input.filled.xlsx"
    output_workbook = load_workbook(result.output_path)
    assert output_workbook.active["C3"].value == "EMPTY"

