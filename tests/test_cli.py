from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook

from excel_fillna.cli import main

FIXTURE_PATH = Path(__file__).with_name("fixtures") / "preserved_artifacts.xlsx"
ARTIFACT_PREFIXES = ("xl/drawings/", "xl/media/", "xl/charts/", "xl/richData/")
ARTIFACT_FILES = {
    "[Content_Types].xml",
    "xl/_rels/workbook.xml.rels",
    "xl/metadata.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/drawings/_rels/drawing1.xml.rels",
}


def _artifact_payloads(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {
            name: archive.read(name)
            for name in sorted(archive.namelist())
            if name.startswith(ARTIFACT_PREFIXES) or name in ARTIFACT_FILES
        }


def _worksheet_cell_xml(
    path: Path,
    coordinate: str,
    worksheet_path: str = "xl/worksheets/sheet1.xml",
) -> tuple[dict[str, str], str | None]:
    with ZipFile(path) as archive:
        sheet_root = ET.fromstring(archive.read(worksheet_path))

    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    cell = sheet_root.find(f".//main:c[@r='{coordinate}']", namespace)
    assert cell is not None
    return dict(cell.attrib), cell.findtext("main:v", None, namespace)


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


def test_cli_preserves_fixture_artifacts(tmp_path: Path, capsys) -> None:
    output = tmp_path / "fixture-output.xlsx"
    source_workbook = load_workbook(FIXTURE_PATH)
    source_worksheet = source_workbook["Data"]
    source_image_count = len(getattr(source_worksheet, "_images", []))
    source_chart_count = len(getattr(source_worksheet, "_charts", []))
    source_workbook.close()

    exit_code = main(
        [
            str(FIXTURE_PATH),
            "--sheet",
            "Data",
            "--range",
            "B1:C1",
            "--output",
            str(output),
            "--fill-text",
            "CLI",
        ]
    )

    assert exit_code == 0

    output_workbook = load_workbook(output)
    output_worksheet = output_workbook["Data"]
    captured = capsys.readouterr()

    assert output_worksheet["B1"].value == "CLI"
    assert output_worksheet["C1"].value == "CLI"
    assert len(getattr(output_worksheet, "_images", [])) == source_image_count
    assert len(getattr(output_worksheet, "_charts", [])) == source_chart_count
    assert _artifact_payloads(output) == _artifact_payloads(FIXTURE_PATH)
    cell_attributes, cell_value = _worksheet_cell_xml(output, "D3")
    assert cell_attributes["t"] == "e"
    assert cell_attributes["vm"] == "1"
    assert cell_value == "#VALUE!"
    assert "filled 2 empty cells" in captured.out
