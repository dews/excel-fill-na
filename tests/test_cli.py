from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_fillna.cli import main


def test_cli_applies_fill_merge_and_output_settings(tmp_path: Path, capsys) -> None:
    source = tmp_path / "input.xlsx"
    output = tmp_path / "output.xlsx"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A4"] = "stop"
    workbook.save(source)

    exit_code = main(
        [
            str(source),
            "--sheet",
            "Data",
            "--range",
            "A1:A4",
            "--merge-empty-runs",
            "--output",
            str(output),
            "--fill-text",
            "MISSING",
        ]
    )

    assert exit_code == 0

    output_workbook = load_workbook(output)
    output_worksheet = output_workbook["Data"]
    merged_ranges = {str(cell_range) for cell_range in output_worksheet.merged_cells.ranges}
    captured = capsys.readouterr()

    assert "A1:A3" in merged_ranges
    assert output_worksheet["A1"].value == "MISSING"
    assert "created 1 merged range" in captured.out
