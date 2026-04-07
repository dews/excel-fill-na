from pathlib import Path
import re
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

FIXTURE_PATH = Path(__file__).with_name("fixtures") / "preserved_artifacts.xlsx"
ARTIFACT_PREFIXES = ("xl/drawings/", "xl/media/", "xl/charts/", "xl/richData/")
ARTIFACT_FILES = {
    "[Content_Types].xml",
    "xl/_rels/workbook.xml.rels",
    "xl/metadata.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/drawings/_rels/drawing1.xml.rels",
}


def artifact_payloads(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {
            name: archive.read(name)
            for name in sorted(archive.namelist())
            if name.startswith(ARTIFACT_PREFIXES) or name in ARTIFACT_FILES
        }


def root_namespace_declarations(path: Path, worksheet_path: str = "xl/worksheets/sheet1.xml") -> set[str]:
    with ZipFile(path) as archive:
        xml_text = archive.read(worksheet_path).decode("utf-8")

    opening_tag_match = re.search(r"<([A-Za-z_][^>\s/]*)\b[^>]*>", xml_text)
    assert opening_tag_match is not None
    opening_tag = opening_tag_match.group(0)
    return set(re.findall(r'xmlns(?::[A-Za-z_][\w.-]*)?="[^"]+"', opening_tag))


def worksheet_cell_xml(
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


def worksheet_hyperlink_refs(
    path: Path,
    worksheet_path: str = "xl/worksheets/sheet1.xml",
) -> list[str]:
    with ZipFile(path) as archive:
        sheet_root = ET.fromstring(archive.read(worksheet_path))

    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    return [
        hyperlink.attrib["ref"]
        for hyperlink in sheet_root.findall("main:hyperlinks/main:hyperlink", namespace)
        if "ref" in hyperlink.attrib
    ]


def worksheet_merge_refs(
    path: Path,
    worksheet_path: str = "xl/worksheets/sheet1.xml",
) -> list[str]:
    with ZipFile(path) as archive:
        sheet_root = ET.fromstring(archive.read(worksheet_path))

    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    return [
        merge_cell.attrib["ref"]
        for merge_cell in sheet_root.findall("main:mergeCells/main:mergeCell", namespace)
        if "ref" in merge_cell.attrib
    ]


def comment_refs(path: Path, comments_path: str = "xl/comments1.xml") -> list[str]:
    with ZipFile(path) as archive:
        comments_root = ET.fromstring(archive.read(comments_path))

    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    return [
        comment.attrib["ref"]
        for comment in comments_root.findall("main:commentList/main:comment", namespace)
        if "ref" in comment.attrib
    ]


def threaded_comment_refs(
    path: Path,
    threaded_comments_path: str = "xl/threadedComments/threadedComment1.xml",
) -> list[str]:
    with ZipFile(path) as archive:
        threaded_root = ET.fromstring(archive.read(threaded_comments_path))

    namespace = {
        "tc": "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
    }
    return [
        comment.attrib["ref"]
        for comment in threaded_root.findall("tc:threadedComment", namespace)
        if "ref" in comment.attrib
    ]


def drawing_anchor_rows(
    path: Path,
    drawing_path: str = "xl/drawings/drawing1.xml",
) -> list[tuple[str, int, int | None]]:
    with ZipFile(path) as archive:
        drawing_root = ET.fromstring(archive.read(drawing_path))

    namespace = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
    rows: list[tuple[str, int, int | None]] = []

    for anchor in drawing_root.findall("xdr:oneCellAnchor", namespace):
        row = anchor.findtext("xdr:from/xdr:row", None, namespace)
        assert row is not None
        rows.append(("oneCell", int(row), None))

    for anchor in drawing_root.findall("xdr:twoCellAnchor", namespace):
        from_row = anchor.findtext("xdr:from/xdr:row", None, namespace)
        to_row = anchor.findtext("xdr:to/xdr:row", None, namespace)
        assert from_row is not None
        assert to_row is not None
        rows.append(("twoCell", int(from_row), int(to_row)))

    return rows
