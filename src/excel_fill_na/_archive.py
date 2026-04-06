from __future__ import annotations

from io import BytesIO
from pathlib import Path
import posixpath
import re
from shutil import copyfile
from tempfile import NamedTemporaryFile
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from openpyxl.worksheet.cell_range import CellRange

from ._models import FillPlan
from ._ranges import parse_range

SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

SHEET_NAMESPACES = {"main": SPREADSHEETML_NS}
WORKBOOK_NAMESPACES = {
    "main": SPREADSHEETML_NS,
    "r": OFFICE_DOCUMENT_REL_NS,
    "rels": PACKAGE_REL_NS,
}


def persist_workbook_changes(
    *,
    source: Path,
    destination: Path,
    worksheet_path: str,
    plan: FillPlan,
) -> None:
    same_path = source.resolve() == destination.resolve()
    if not plan.cell_writes and not plan.merged_ranges:
        if same_path:
            return
        copyfile(source, destination)
        return

    if same_path:
        write_patched_archive_in_place(
            source=source,
            destination=destination,
            worksheet_path=worksheet_path,
            plan=plan,
        )
        return

    write_patched_archive(
        source=source,
        destination=destination,
        worksheet_path=worksheet_path,
        plan=plan,
    )


def resolve_worksheet_archive_path(source: Path, sheet_name: str) -> str:
    with ZipFile(source) as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    relationship_targets = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in relationships_root.findall(package_relationship_tag("Relationship"))
    }

    sheets = workbook_root.find("main:sheets", WORKBOOK_NAMESPACES)
    if sheets is None:
        raise ValueError("Workbook metadata is missing sheet definitions.")

    relationship_attr = f"{{{OFFICE_DOCUMENT_REL_NS}}}id"
    for sheet in sheets.findall("main:sheet", WORKBOOK_NAMESPACES):
        if sheet.attrib.get("name") != sheet_name:
            continue

        relationship_id = sheet.attrib.get(relationship_attr)
        if relationship_id is None or relationship_id not in relationship_targets:
            raise ValueError(f"Worksheet {sheet_name!r} is missing workbook relationship metadata.")
        return normalize_archive_path("xl/workbook.xml", relationship_targets[relationship_id])

    raise ValueError(f"Worksheet {sheet_name!r} was not found in workbook metadata.")


def find_value_metadata_cells(source: Path, worksheet_path: str) -> set[tuple[int, int]]:
    with ZipFile(source) as archive:
        worksheet_root = ET.fromstring(archive.read(worksheet_path))

    sheet_data = worksheet_root.find("main:sheetData", SHEET_NAMESPACES)
    if sheet_data is None:
        return set()

    coordinates: set[tuple[int, int]] = set()
    for row_element in sheet_data.findall(sheet_tag("row")):
        for cell_element in row_element.findall(sheet_tag("c")):
            coordinate = cell_element.attrib.get("r")
            if coordinate is None or "vm" not in cell_element.attrib:
                continue
            coordinates.add(
                (
                    coordinate_row_index(coordinate),
                    coordinate_column_index(coordinate),
                )
            )

    return coordinates


def write_patched_archive_in_place(
    *,
    source: Path,
    destination: Path,
    worksheet_path: str,
    plan: FillPlan,
) -> None:
    temporary_path: Path | None = None
    try:
        with NamedTemporaryFile(
            prefix=f"{source.stem}.",
            suffix=source.suffix,
            dir=source.parent,
            delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
        write_patched_archive(
            source=source,
            destination=temporary_path,
            worksheet_path=worksheet_path,
            plan=plan,
        )
        temporary_path.replace(destination)
    except Exception:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()
        raise


def write_patched_archive(
    *,
    source: Path,
    destination: Path,
    worksheet_path: str,
    plan: FillPlan,
) -> None:
    with ZipFile(source) as source_archive:
        worksheet_xml = source_archive.read(worksheet_path)
        patched_worksheet_xml = patch_worksheet_xml(worksheet_xml, plan)

        with ZipFile(destination, "w") as destination_archive:
            for info in source_archive.infolist():
                data = patched_worksheet_xml if info.filename == worksheet_path else source_archive.read(
                    info.filename
                )
                destination_archive.writestr(info, data)


def patch_worksheet_xml(worksheet_xml: bytes, plan: FillPlan) -> bytes:
    worksheet_root, namespaces = parse_xml_bytes(worksheet_xml)
    register_namespaces(namespaces)

    sheet_data = worksheet_root.find("main:sheetData", SHEET_NAMESPACES)
    if sheet_data is None:
        raise ValueError("Worksheet XML is missing sheetData.")

    row_lookup = {
        safe_int(row.attrib.get("r")): row
        for row in sheet_data.findall(sheet_tag("row"))
        if safe_int(row.attrib.get("r")) is not None
    }
    touched_rows: set[int] = set()

    for cell_write in plan.cell_writes:
        row_element = get_or_create_row(sheet_data, row_lookup, cell_write.row)
        cell_element = get_or_create_cell(row_element, cell_write.row, cell_write.column)
        set_inline_string_value(cell_element, cell_write.value)
        touched_rows.add(cell_write.row)

    for row_index in touched_rows:
        update_row_spans(row_lookup[row_index])

    if plan.merged_ranges:
        append_merge_ranges(worksheet_root, plan.merged_ranges)

    update_dimension(worksheet_root, sheet_data, plan)

    patched_xml = ET.tostring(worksheet_root, encoding="utf-8", xml_declaration=True)
    return restore_root_namespace_declarations(patched_xml, namespaces)


def normalize_archive_path(base_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_path), target))


def parse_xml_bytes(xml_bytes: bytes) -> tuple[ET.Element, list[tuple[str, str]]]:
    namespaces: list[tuple[str, str]] = []
    seen_namespaces: set[tuple[str, str]] = set()
    for _, namespace in ET.iterparse(BytesIO(xml_bytes), events=("start-ns",)):
        if namespace not in seen_namespaces:
            namespaces.append(namespace)
            seen_namespaces.add(namespace)
    return ET.fromstring(xml_bytes), namespaces


def register_namespaces(namespaces: Iterable[tuple[str, str]]) -> None:
    for prefix, uri in namespaces:
        ET.register_namespace(prefix, uri)


def restore_root_namespace_declarations(
    xml_bytes: bytes,
    namespaces: Iterable[tuple[str, str]],
) -> bytes:
    xml_text = xml_bytes.decode("utf-8")
    root_match = re.search(r"<([A-Za-z_][^>\s/]*)\b[^>]*>", xml_text)
    if root_match is None:
        return xml_bytes

    opening_tag = root_match.group(0)
    declarations: list[str] = []
    for prefix, uri in namespaces:
        declaration = f'xmlns:{prefix}="{uri}"' if prefix else f'xmlns="{uri}"'
        if declaration not in opening_tag:
            declarations.append(declaration)

    if not declarations:
        return xml_bytes

    insertion = " " + " ".join(declarations)
    updated_text = xml_text.replace(opening_tag, opening_tag[:-1] + insertion + ">", 1)
    return updated_text.encode("utf-8")


def get_or_create_row(
    sheet_data: ET.Element,
    row_lookup: dict[int, ET.Element],
    row_index: int,
) -> ET.Element:
    existing_row = row_lookup.get(row_index)
    if existing_row is not None:
        return existing_row

    row_element = ET.Element(sheet_tag("row"), {"r": str(row_index)})
    insert_at = len(sheet_data)
    for position, candidate in enumerate(sheet_data):
        candidate_row = safe_int(candidate.attrib.get("r"))
        if candidate_row is not None and candidate_row > row_index:
            insert_at = position
            break
    sheet_data.insert(insert_at, row_element)
    row_lookup[row_index] = row_element
    return row_element


def get_or_create_cell(row_element: ET.Element, row_index: int, column_index: int) -> ET.Element:
    coordinate = f"{get_column_letter(column_index)}{row_index}"
    insert_at = len(row_element)

    for position, candidate in enumerate(row_element):
        if candidate.tag != sheet_tag("c"):
            continue

        candidate_coordinate = candidate.attrib.get("r")
        if candidate_coordinate == coordinate:
            return candidate
        if candidate_coordinate is not None and coordinate_column_index(candidate_coordinate) > column_index:
            insert_at = position
            break

    cell_element = ET.Element(sheet_tag("c"), {"r": coordinate})
    row_element.insert(insert_at, cell_element)
    return cell_element


def set_inline_string_value(cell_element: ET.Element, value: str) -> None:
    preserved_attributes = {
        name: attribute_value
        for name, attribute_value in cell_element.attrib.items()
        if name not in ("t", "vm")
    }

    for child in list(cell_element):
        cell_element.remove(child)

    cell_element.attrib.clear()
    cell_element.attrib.update(preserved_attributes)
    cell_element.set("t", "inlineStr")

    inline_string = ET.SubElement(cell_element, sheet_tag("is"))
    text_element = ET.SubElement(inline_string, sheet_tag("t"))
    if value != value.strip():
        text_element.set(XML_SPACE, "preserve")
    text_element.text = value


def append_merge_ranges(worksheet_root: ET.Element, merged_ranges: Iterable[str]) -> None:
    merge_cells = worksheet_root.find("main:mergeCells", SHEET_NAMESPACES)
    if merge_cells is None:
        merge_cells = ET.Element(sheet_tag("mergeCells"))
        insert_merge_cells_element(worksheet_root, merge_cells)

    existing_ranges = {
        merge_cell.attrib.get("ref")
        for merge_cell in merge_cells.findall(sheet_tag("mergeCell"))
    }
    for merged_range in merged_ranges:
        if merged_range not in existing_ranges:
            ET.SubElement(merge_cells, sheet_tag("mergeCell"), {"ref": merged_range})
            existing_ranges.add(merged_range)

    merge_cells.set(
        "count",
        str(sum(1 for child in merge_cells if child.tag == sheet_tag("mergeCell"))),
    )


def insert_merge_cells_element(worksheet_root: ET.Element, merge_cells: ET.Element) -> None:
    preceding_names = {
        "sheetData",
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
    }
    insert_at: int | None = None
    for index, child in enumerate(list(worksheet_root)):
        if local_name(child.tag) in preceding_names:
            insert_at = index + 1

    if insert_at is None:
        worksheet_root.append(merge_cells)
        return

    worksheet_root.insert(insert_at, merge_cells)


def update_dimension(worksheet_root: ET.Element, sheet_data: ET.Element, plan: FillPlan) -> None:
    bounds = existing_dimension_bounds(worksheet_root)
    if bounds is None:
        bounds = scan_sheet_data_bounds(sheet_data)

    for cell_write in plan.cell_writes:
        bounds = expand_bounds(bounds, row=cell_write.row, column=cell_write.column)
    for merged_range in plan.merged_ranges:
        bounds = expand_bounds_with_range(bounds, parse_range(merged_range))

    if bounds is None:
        return

    ref = format_dimension(bounds)
    dimension = worksheet_root.find("main:dimension", SHEET_NAMESPACES)
    if dimension is None:
        dimension = ET.Element(sheet_tag("dimension"), {"ref": ref})
        insert_dimension_element(worksheet_root, dimension)
        return

    dimension.set("ref", ref)


def existing_dimension_bounds(worksheet_root: ET.Element) -> tuple[int, int, int, int] | None:
    dimension = worksheet_root.find("main:dimension", SHEET_NAMESPACES)
    if dimension is None:
        return None

    ref = dimension.attrib.get("ref")
    if not ref:
        return None

    try:
        cell_range = parse_range(ref)
    except ValueError:
        return None

    return (
        cell_range.min_row,
        cell_range.min_col,
        cell_range.max_row,
        cell_range.max_col,
    )


def scan_sheet_data_bounds(sheet_data: ET.Element) -> tuple[int, int, int, int] | None:
    bounds: tuple[int, int, int, int] | None = None
    for row_element in sheet_data.findall(sheet_tag("row")):
        row_index = safe_int(row_element.attrib.get("r"))
        if row_index is None:
            continue

        for cell_element in row_element.findall(sheet_tag("c")):
            coordinate = cell_element.attrib.get("r")
            if coordinate is None:
                continue
            bounds = expand_bounds(
                bounds,
                row=row_index,
                column=coordinate_column_index(coordinate),
            )
    return bounds


def insert_dimension_element(worksheet_root: ET.Element, dimension: ET.Element) -> None:
    insert_at = 0
    for index, child in enumerate(list(worksheet_root)):
        if local_name(child.tag) == "sheetPr":
            insert_at = index + 1
            break
        insert_at = index
        break
    worksheet_root.insert(insert_at, dimension)


def expand_bounds(
    bounds: tuple[int, int, int, int] | None,
    *,
    row: int,
    column: int,
) -> tuple[int, int, int, int]:
    if bounds is None:
        return (row, column, row, column)

    min_row, min_col, max_row, max_col = bounds
    return (
        min(min_row, row),
        min(min_col, column),
        max(max_row, row),
        max(max_col, column),
    )


def expand_bounds_with_range(
    bounds: tuple[int, int, int, int] | None,
    cell_range: CellRange,
) -> tuple[int, int, int, int]:
    bounds = expand_bounds(bounds, row=cell_range.min_row, column=cell_range.min_col)
    return expand_bounds(bounds, row=cell_range.max_row, column=cell_range.max_col)


def format_dimension(bounds: tuple[int, int, int, int]) -> str:
    min_row, min_col, max_row, max_col = bounds
    start = f"{get_column_letter(min_col)}{min_row}"
    end = f"{get_column_letter(max_col)}{max_row}"
    return start if start == end else f"{start}:{end}"


def update_row_spans(row_element: ET.Element) -> None:
    columns = [
        coordinate_column_index(cell.attrib["r"])
        for cell in row_element
        if cell.tag == sheet_tag("c") and "r" in cell.attrib
    ]
    if not columns:
        row_element.attrib.pop("spans", None)
        return

    row_element.set("spans", f"{min(columns)}:{max(columns)}")


def coordinate_column_index(coordinate: str) -> int:
    column_letters, _ = coordinate_from_string(coordinate)
    return column_index_from_string(column_letters)


def coordinate_row_index(coordinate: str) -> int:
    _, row_index = coordinate_from_string(coordinate)
    return row_index


def sheet_tag(local_name: str) -> str:
    return f"{{{SPREADSHEETML_NS}}}{local_name}"


def package_relationship_tag(local_name: str) -> str:
    return f"{{{PACKAGE_REL_NS}}}{local_name}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def safe_int(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None
