"""Read and write .xlsx workbooks with the Python 3.9 standard library."""

from __future__ import annotations

import os
import re
import zipfile
from typing import Dict, List, Optional, Sequence
from xml.etree import ElementTree as ET

from jmeter_gen.model import InputError, Sheet, xml_escape

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_BAD_SHEET = re.compile(r"[:\\/?*\[\]]")


def _local(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def _col_index(ref: str) -> int:
    letters = "".join(ch for ch in ref if ch.isalpha())
    if not letters:
        raise InputError(f"セル参照が不正です: {ref}")
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch.upper()) - 64)
    return index


def _col_letter(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _number_to_text(raw: str) -> str:
    text = raw.strip()
    if re.fullmatch(r"-?\d+", text):
        return text
    if re.fullmatch(r"-?\d+\.\d+(?:[eE][+-]?\d+)?", text) or re.fullmatch(
        r"-?\d+[eE][+-]?\d+", text
    ):
        try:
            value = float(text)
        except ValueError:
            return text
        if value.is_integer() and abs(value) < 1e15:
            return str(int(value))
    return text


def _collect_text(node: ET.Element) -> str:
    parts: List[str] = []
    if _local(node.tag) == "t" and node.text:
        parts.append(node.text)
    for child in list(node):
        if _local(child.tag) in ("rPh", "phoneticPr"):
            continue
        parts.append(_collect_text(child))
    return "".join(parts)


def _cell_text(cell: ET.Element, shared: Sequence[str]) -> str:
    kind = cell.attrib.get("t", "")
    if kind == "inlineStr":
        return _collect_text(cell)
    value_node = None
    for child in list(cell):
        if _local(child.tag) == "v":
            value_node = child
            break
    if kind == "s":
        if value_node is None or value_node.text is None:
            return ""
        try:
            return shared[int(value_node.text)]
        except (ValueError, IndexError) as exc:
            raise InputError("Excelの共有文字列を読めません。") from exc
    if kind == "b":
        if value_node is None or value_node.text is None:
            return ""
        return "true" if value_node.text.strip() in ("1", "true") else "false"
    if kind in ("str", "n", ""):
        if value_node is None or value_node.text is None:
            if kind == "str":
                return _collect_text(cell)
            return ""
        if kind == "n" or kind == "":
            return _number_to_text(value_node.text)
        return value_node.text
    if value_node is not None and value_node.text:
        return value_node.text
    return ""


def _sheet_rows(root: ET.Element, shared: Sequence[str]) -> List[tuple]:
    sheet_data = None
    for child in list(root):
        if _local(child.tag) == "sheetData":
            sheet_data = child
            break
    if sheet_data is None:
        return []
    rows = []
    next_number = 1
    for row in list(sheet_data):
        if _local(row.tag) != "row":
            continue
        raw_number = row.attrib.get("r", "")
        if raw_number.isdigit():
            number = int(raw_number)
        else:
            number = next_number
        next_number = number + 1
        cells: Dict[int, str] = {}
        max_col = 0
        sequential = 1
        for cell in list(row):
            if _local(cell.tag) != "c":
                continue
            ref = cell.attrib.get("r", "")
            if ref:
                column = _col_index(ref)
            else:
                column = sequential
            sequential = column + 1
            cells[column] = _cell_text(cell, shared)
            if column > max_col:
                max_col = column
        values = [] if max_col == 0 else [cells.get(i, "") for i in range(1, max_col + 1)]
        rows.append((number, values))
    return rows


def _read_shared_strings(blob: bytes) -> List[str]:
    root = ET.fromstring(blob)
    values: List[str] = []
    for child in list(root):
        if _local(child.tag) == "si":
            values.append(_collect_text(child))
    return values


def _resolve_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        parts = target.lstrip("/").split("/")
    else:
        parts = [part for part in base_dir.split("/") if part] + target.split("/")
    out: List[str] = []
    for part in parts:
        if part in ("", "."):
            continue
        if part == "..":
            if out:
                out.pop()
            continue
        out.append(part)
    return "/".join(out)


def read_xlsx(path: str) -> List[Sheet]:
    if not os.path.isfile(path):
        raise InputError(f"入力ファイルがありません: {path}")
    try:
        archive = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise InputError(f"Excelファイルとして読めません: {path}") from exc
    try:
        names = set(archive.namelist())
        if "xl/workbook.xml" not in names:
            raise InputError(f"workbook.xml がありません: {path}")
        shared: List[str] = []
        if "xl/sharedStrings.xml" in names:
            shared = _read_shared_strings(archive.read("xl/sharedStrings.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map: Dict[str, str] = {}
        for rel in list(rels_root):
            if _local(rel.tag) != "Relationship":
                continue
            rel_id = rel.attrib.get("Id", "")
            target = rel.attrib.get("Target", "")
            rel_type = rel.attrib.get("Type", "")
            if rel_id and "worksheet" in rel_type:
                rel_map[rel_id] = _resolve_target("xl", target)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets: List[Sheet] = []
        for child in workbook.iter():
            if _local(child.tag) != "sheet":
                continue
            name = child.attrib.get("name", "")
            rel_id = child.attrib.get(f"{{{_REL}}}id", "")
            if not rel_id:
                for key, value in child.attrib.items():
                    if _local(key) == "id":
                        rel_id = value
                        break
            target = rel_map.get(rel_id, "")
            if not target or target not in names:
                raise InputError(f"シートの実体がありません: {name}")
            numbered = _sheet_rows(ET.fromstring(archive.read(target)), shared)
            sheets.append(
                Sheet(
                    name=name,
                    rows=[values for _, values in numbered],
                    row_numbers=[number for number, _ in numbered],
                )
            )
        if not sheets:
            raise InputError("Excelにシートがありません。")
        return sheets
    finally:
        archive.close()


def _safe_sheet_name(name: str, used: set) -> str:
    cleaned = _BAD_SHEET.sub("_", name).strip().strip("'") or "Sheet"
    cleaned = cleaned[:31]
    base = cleaned
    number = 2
    while cleaned.lower() in used:
        suffix = f"_{number}"
        cleaned = base[: 31 - len(suffix)] + suffix
        number += 1
    used.add(cleaned.lower())
    return cleaned


def _worksheet_xml(sheet: Sheet) -> str:
    rows = sheet.rows
    width = 1
    for row in rows:
        if len(row) > width:
            width = len(row)
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{_MAIN}">',
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        "</sheetView></sheetViews>",
        "<cols>",
    ]
    widths = list(sheet.widths or [])
    for index in range(1, width + 1):
        size = widths[index - 1] if index - 1 < len(widths) else 24
        lines.append(
            f'<col min="{index}" max="{index}" width="{size}" customWidth="1"/>'
        )
    lines.append("</cols>")
    lines.append("<sheetData>")
    wrap_cols = set(sheet.wrap_cols or [])
    for row_number, row in enumerate(rows, start=1):
        lines.append(f'<row r="{row_number}">')
        for col_number, value in enumerate(row, start=1):
            if value is None or value == "":
                continue
            style = "0"
            if row_number == 1:
                style = "1"
            elif col_number in wrap_cols:
                style = "2"
            ref = f"{_col_letter(col_number)}{row_number}"
            lines.append(
                f'<c r="{ref}" t="inlineStr" s="{style}">'
                f'<is><t xml:space="preserve">{xml_escape(value)}</t></is></c>'
            )
        lines.append("</row>")
    lines.append("</sheetData>")
    if rows:
        ref = f"A1:{_col_letter(width)}{len(rows)}"
        lines.append(f'<autoFilter ref="{ref}"/>')
    lines.append("</worksheet>")
    return "\n".join(lines) + "\n"


def _content_types(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    body = "\n  ".join(overrides)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Types xmlns="{_CT}">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        f"  {body}\n"
        "</Types>\n"
    )


def _styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<styleSheet xmlns="{_MAIN}">\n'
        '  <fonts count="2">\n'
        '    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>\n'
        '    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>\n'
        "  </fonts>\n"
        '  <fills count="3">\n'
        '    <fill><patternFill patternType="none"/></fill>\n'
        '    <fill><patternFill patternType="gray125"/></fill>\n'
        '    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/>'
        "<bgColor rgb=\"FF1F4E79\"/></patternFill></fill>\n"
        "  </fills>\n"
        '  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n'
        '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'
        '  <cellXfs count="3">\n'
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n'
        '    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">'
        '<alignment wrapText="1" vertical="center"/></xf>\n'
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
        '<alignment wrapText="1" vertical="top"/></xf>\n'
        "  </cellXfs>\n"
        "</styleSheet>\n"
    )


def write_xlsx(path: str, sheets: Sequence[Sheet]) -> None:
    if not sheets:
        raise InputError("書き出すシートがありません。")
    directory = os.path.dirname(os.path.abspath(path))
    if directory:
        os.makedirs(directory, exist_ok=True)
    used: set = set()
    safe_names = [_safe_sheet_name(sheet.name, used) for sheet in sheets]
    workbook_sheets = []
    workbook_rels = [
        f'<Relationship Id="rIdStyles" Type="{_REL}/styles" Target="styles.xml"/>'
    ]
    for index, name in enumerate(safe_names, start=1):
        workbook_sheets.append(
            f'<sheet name="{xml_escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        )
        workbook_rels.append(
            f'<Relationship Id="rId{index}" Type="{_REL}/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<workbook xmlns="{_MAIN}" xmlns:r="{_REL}">\n'
        "  <sheets>\n    "
        + "\n    ".join(workbook_sheets)
        + "\n  </sheets>\n</workbook>\n"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{_PKG}">\n  '
        + "\n  ".join(workbook_rels)
        + "\n</Relationships>\n"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{_PKG}">\n'
        f'  <Relationship Id="rId1" Type="{_REL}/officeDocument" Target="xl/workbook.xml"/>\n'
        "</Relationships>\n"
    )
    temp_path = path + ".tmp"
    if os.path.exists(temp_path):
        os.remove(temp_path)
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        def add(name: str, text: str) -> None:
            archive.writestr(name, text.encode("utf-8"))

        add("[Content_Types].xml", _content_types(len(sheets)))
        add("_rels/.rels", root_rels)
        add("xl/workbook.xml", workbook)
        add("xl/_rels/workbook.xml.rels", rels)
        add("xl/styles.xml", _styles_xml())
        for index, sheet in enumerate(sheets, start=1):
            add(f"xl/worksheets/sheet{index}.xml", _worksheet_xml(sheet))
    os.replace(temp_path, path)
