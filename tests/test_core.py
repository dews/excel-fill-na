from pathlib import Path
import re
from zipfile import ZipFile
from zipfile import ZIP_DEFLATED
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment

from excel_fillna.core import fill_empty_cells, process_workbook

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


def _root_namespace_declarations(path: Path, worksheet_path: str = "xl/worksheets/sheet1.xml") -> set[str]:
    with ZipFile(path) as archive:
        xml_text = archive.read(worksheet_path).decode("utf-8")

    opening_tag_match = re.search(r"<([A-Za-z_][^>\s/]*)\b[^>]*>", xml_text)
    assert opening_tag_match is not None
    opening_tag = opening_tag_match.group(0)
    return set(re.findall(r'xmlns(?::[A-Za-z_][\w.-]*)?="[^"]+"', opening_tag))


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


def _write_comment_only_workbook(path: Path) -> None:
    base = path.with_name("comment-only.base.xlsx")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = None
    worksheet["A2"] = "placeholder"
    worksheet["A2"].comment = Comment("note", "tester")
    worksheet["A3"] = "stop"
    workbook.save(base)
    workbook.close()

    with ZipFile(base) as source_archive:
        sheet_root = ET.fromstring(source_archive.read("xl/worksheets/sheet1.xml"))
        comments_root = ET.fromstring(source_archive.read("xl/comments/comment1.xml"))

        namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        row = sheet_root.find("main:sheetData/main:row[@r='2']", namespace)
        assert row is not None

        cell = row.find("main:c[@r='A2']", namespace)
        assert cell is not None
        for child in list(cell):
            cell.remove(child)
        cell.attrib.clear()
        cell.set("r", "A2")

        comment = comments_root.find("main:commentList/main:comment", namespace)
        assert comment is not None
        comment.set("ref", "A2")

        with ZipFile(path, "w", compression=ZIP_DEFLATED) as destination_archive:
            for info in source_archive.infolist():
                data = source_archive.read(info.filename)
                if info.filename == "xl/worksheets/sheet1.xml":
                    data = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
                elif info.filename == "xl/comments/comment1.xml":
                    data = ET.tostring(comments_root, encoding="utf-8", xml_declaration=True)
                destination_archive.writestr(info, data)

    base.unlink()


def _write_value_metadata_only_workbook(path: Path) -> None:
    base = path.with_name("value-metadata-only.base.xlsx")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = None
    worksheet["A2"] = "placeholder"
    worksheet["A3"] = "stop"
    workbook.save(base)
    workbook.close()

    with ZipFile(base) as source_archive:
        sheet_root = ET.fromstring(source_archive.read("xl/worksheets/sheet1.xml"))

        namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        row = sheet_root.find("main:sheetData/main:row[@r='2']", namespace)
        assert row is not None

        cell = row.find("main:c[@r='A2']", namespace)
        assert cell is not None
        for child in list(cell):
            cell.remove(child)
        cell.attrib.clear()
        cell.set("r", "A2")
        cell.set("vm", "1")

        with ZipFile(path, "w", compression=ZIP_DEFLATED) as destination_archive:
            for info in source_archive.infolist():
                data = source_archive.read(info.filename)
                if info.filename == "xl/worksheets/sheet1.xml":
                    data = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
                destination_archive.writestr(info, data)

    base.unlink()


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


def test_fill_empty_cells_keeps_comment_only_cells_unfilled_and_unmerged() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet["A2"].comment = Comment("note", "tester")
    worksheet["A3"] = "stop"

    result = fill_empty_cells(
        worksheet,
        target_range="A1:A3",
        merge_empty_runs=True,
    )

    assert result.filled_cells == 1
    assert result.merged_ranges == ()
    assert not worksheet.merged_cells.ranges
    assert worksheet["A1"].value == "NA"
    assert worksheet["A2"].value is None
    assert worksheet["A2"].comment is not None
    assert worksheet["A2"].comment.text == "note"
    assert worksheet["A3"].value == "stop"


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


def test_process_workbook_preserves_fixture_artifacts_and_styles(tmp_path: Path) -> None:
    output = tmp_path / "preserved.xlsx"

    source_workbook = load_workbook(FIXTURE_PATH)
    source_worksheet = source_workbook["Data"]
    source_style_id = source_worksheet["C1"].style_id
    source_image_count = len(getattr(source_worksheet, "_images", []))
    source_chart_count = len(getattr(source_worksheet, "_charts", []))
    source_workbook.close()

    result = process_workbook(
        FIXTURE_PATH,
        sheet_name="Data",
        target_range="B1:C1",
        fill_value="EMPTY",
        output_path=output,
    )

    assert result.filled_cells == 2

    output_workbook = load_workbook(output)
    output_worksheet = output_workbook["Data"]

    assert output_worksheet["B1"].value == "EMPTY"
    assert output_worksheet["C1"].value == "EMPTY"
    assert output_worksheet["C1"].style_id == source_style_id
    assert len(getattr(output_worksheet, "_images", [])) == source_image_count
    assert len(getattr(output_worksheet, "_charts", [])) == source_chart_count

    artifact_payloads = _artifact_payloads(output)
    assert "xl/charts/chart1.xml" in artifact_payloads
    assert "xl/richData/richValueRel.xml" in artifact_payloads
    assert "xl/worksheets/_rels/sheet1.xml.rels" in artifact_payloads
    assert artifact_payloads == _artifact_payloads(FIXTURE_PATH)
    assert _root_namespace_declarations(output) == _root_namespace_declarations(FIXTURE_PATH)
    cell_attributes, cell_value = _worksheet_cell_xml(output, "D3")
    assert cell_attributes["t"] == "e"
    assert cell_attributes["vm"] == "1"
    assert cell_value == "#VALUE!"


def test_process_workbook_targets_named_sheet_without_touching_other_sheets(tmp_path: Path) -> None:
    output = tmp_path / "other-sheet.xlsx"

    result = process_workbook(
        FIXTURE_PATH,
        sheet_name="Other",
        target_range="B2",
        fill_value="FILLED",
        output_path=output,
    )

    assert result.sheet_name == "Other"
    assert result.filled_cells == 1

    output_workbook = load_workbook(output)
    assert output_workbook["Other"]["A1"].value == "keep"
    assert output_workbook["Other"]["B2"].value == "FILLED"
    assert output_workbook["Data"]["B1"].value is None
    assert output_workbook["Data"]["C1"].value is None
    assert _artifact_payloads(output) == _artifact_payloads(FIXTURE_PATH)


def test_process_workbook_does_not_merge_comment_only_cells(tmp_path: Path) -> None:
    source = tmp_path / "comment-only.xlsx"
    output = tmp_path / "comment-only.output.xlsx"
    _write_comment_only_workbook(source)

    result = process_workbook(
        source,
        sheet_name="Data",
        target_range="A1:A3",
        merge_empty_runs=True,
        output_path=output,
    )

    output_workbook = load_workbook(output)
    output_worksheet = output_workbook["Data"]

    assert result.filled_cells == 1
    assert result.merged_ranges == ()
    assert not output_worksheet.merged_cells.ranges
    assert output_worksheet["A1"].value == "NA"
    assert output_worksheet["A2"].comment is not None
    assert output_worksheet["A2"].comment.text == "note"


def test_process_workbook_does_not_fill_value_metadata_cells(tmp_path: Path) -> None:
    source = tmp_path / "value-metadata-only.xlsx"
    output = tmp_path / "value-metadata-only.output.xlsx"
    _write_value_metadata_only_workbook(source)

    result = process_workbook(
        source,
        sheet_name="Data",
        target_range="A1:A3",
        merge_empty_runs=True,
        output_path=output,
    )

    output_workbook = load_workbook(output)
    output_worksheet = output_workbook["Data"]
    cell_attributes, cell_value = _worksheet_cell_xml(output, "A2")

    assert result.filled_cells == 1
    assert result.merged_ranges == ()
    assert not output_worksheet.merged_cells.ranges
    assert output_worksheet["A1"].value == "NA"
    assert output_worksheet["A2"].value is None
    assert cell_attributes["vm"] == "1"
    assert cell_value is None
    assert output_worksheet["A3"].value == "stop"
