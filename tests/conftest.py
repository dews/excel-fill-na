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

