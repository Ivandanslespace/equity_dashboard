"""Écriture ciblée de classeurs XLSX sans automatiser Microsoft Excel.

Le moteur conserve toutes les parties du modèle qui ne sont pas modifiées et
n'édite que les feuilles, les styles et les métadonnées de calcul nécessaires.
"""

from __future__ import annotations

import copy
import datetime as dt
import math
import os
import posixpath
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from lxml import etree


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

REL_DRAWING = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
REL_IMAGE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)
REL_CALC_CHAIN = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
)

_A1_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.IGNORECASE)


def _qname(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def _column_to_number(column: str) -> int:
    number = 0
    for char in column.upper():
        number = number * 26 + ord(char) - 64
    return number


def _number_to_column(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _parse_address(address: str) -> tuple[int, int]:
    match = _A1_RE.match(address.strip())
    if not match:
        raise ValueError(f"Adresse de cellule invalide : {address}")
    return int(match.group(2)), _column_to_number(match.group(1))


def _cell_address(row: int, column: int, absolute: bool = False) -> str:
    marker = "$" if absolute else ""
    return f"{marker}{_number_to_column(column)}{marker}{row}"


def _normalise_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _relationship_part(part: str) -> str:
    directory, filename = posixpath.split(part)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _serialise(root: etree._Element) -> bytes:
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value == "":
        return True
    try:
        result = pd.isna(value)
        return bool(result) if not hasattr(result, "__len__") else False
    except (TypeError, ValueError):
        return False


def _python_value(value: Any) -> Any:
    if hasattr(value, "item") and not isinstance(value, (str, bytes)):
        try:
            return value.item()
        except (TypeError, ValueError):
            pass
    return value


class _StyleManager:
    """Crée uniquement les variantes de styles réellement demandées."""

    def __init__(self, workbook: "OoxmlWorkbook") -> None:
        self.workbook = workbook
        self.part = "xl/styles.xml"
        self.root = etree.fromstring(workbook._parts[self.part])
        self.cell_xfs = self.root.find(_qname(NS_MAIN, "cellXfs"))
        if self.cell_xfs is None:
            raise ValueError("La table cellXfs est absente de styles.xml.")
        self.cache: dict[tuple[int, str | None, str | None], int] = {}
        self.dirty = False

    def variant(
        self,
        base_style: int,
        number_format: str | None = None,
        horizontal: str | None = None,
    ) -> int:
        key = (base_style, number_format, horizontal)
        if key in self.cache:
            return self.cache[key]

        styles = list(self.cell_xfs)
        if not 0 <= base_style < len(styles):
            base_style = 0
        xf = copy.deepcopy(styles[base_style])

        if number_format:
            format_id = {"0.00%": 10, "yyyy-mm-dd": 14}.get(number_format)
            if format_id is None:
                format_id = self._custom_number_format(number_format)
            xf.set("numFmtId", str(format_id))
            xf.set("applyNumberFormat", "1")

        if horizontal:
            alignment = xf.find(_qname(NS_MAIN, "alignment"))
            if alignment is None:
                alignment = etree.SubElement(xf, _qname(NS_MAIN, "alignment"))
            alignment.set("horizontal", horizontal)
            xf.set("applyAlignment", "1")

        self.cell_xfs.append(xf)
        self.cell_xfs.set("count", str(len(self.cell_xfs)))
        style_id = len(self.cell_xfs) - 1
        self.cache[key] = style_id
        self.dirty = True
        return style_id

    def _custom_number_format(self, code: str) -> int:
        num_fmts = self.root.find(_qname(NS_MAIN, "numFmts"))
        if num_fmts is None:
            num_fmts = etree.Element(_qname(NS_MAIN, "numFmts"), count="0")
            self.root.insert(0, num_fmts)
        for item in num_fmts:
            if item.get("formatCode") == code:
                return int(item.get("numFmtId"))
        used = {int(item.get("numFmtId")) for item in num_fmts}
        format_id = max(used | {163}) + 1
        etree.SubElement(
            num_fmts,
            _qname(NS_MAIN, "numFmt"),
            numFmtId=str(format_id),
            formatCode=code,
        )
        num_fmts.set("count", str(len(num_fmts)))
        self.dirty = True
        return format_id

    def flush(self) -> None:
        if self.dirty:
            self.workbook._parts[self.part] = _serialise(self.root)


@dataclass
class _Picture:
    collection: "_Pictures"
    name: str
    anchor: etree._Element

    def delete(self) -> None:
        self.anchor.getparent().remove(self.anchor)
        self.collection._dirty = True


class _Pictures:
    """Collection minimale d'images compatible avec les exports existants."""

    def __init__(self, worksheet: "OoxmlWorksheet") -> None:
        self.worksheet = worksheet
        self._dirty = False
        self._drawing_part: str | None = None
        self._drawing_root: etree._Element | None = None
        self._drawing_rels_part: str | None = None
        self._drawing_rels_root: etree._Element | None = None

    def __iter__(self) -> Iterable[_Picture]:
        self._load(create=False)
        if self._drawing_root is None:
            return iter(())
        pictures = []
        for anchor in self._drawing_root:
            name_nodes = anchor.xpath(".//xdr:cNvPr", namespaces={"xdr": NS_XDR})
            if name_nodes:
                pictures.append(_Picture(self, name_nodes[0].get("name", ""), anchor))
        return iter(pictures)

    def add(
        self,
        image_path: str,
        name: str,
        left: float = 0,
        top: float = 0,
        scale: float = 1.0,
    ) -> None:
        del left, top
        self._load(create=True)
        assert self._drawing_root is not None
        assert self._drawing_rels_root is not None

        path = Path(image_path)
        extension = path.suffix.lower().lstrip(".")
        if extension == "jpg":
            content_type = "image/jpeg"
        elif extension == "jpeg":
            content_type = "image/jpeg"
        elif extension == "gif":
            content_type = "image/gif"
        else:
            extension = "png"
            content_type = "image/png"

        media_part = self.worksheet.workbook._next_part(
            "xl/media/image", f".{extension}"
        )
        self.worksheet.workbook._parts[media_part] = path.read_bytes()
        self.worksheet.workbook._new_parts.add(media_part)
        self.worksheet.workbook._ensure_default_content_type(extension, content_type)

        rel_id = self.worksheet.workbook._next_relationship_id(
            self._drawing_rels_root
        )
        etree.SubElement(
            self._drawing_rels_root,
            _qname(NS_PKG_REL, "Relationship"),
            Id=rel_id,
            Type=REL_IMAGE,
            Target=f"../media/{posixpath.basename(media_part)}",
        )

        width, height = self._image_size(path)
        picture_id = max(
            [
                int(node.get("id", "0"))
                for node in self._drawing_root.xpath(
                    ".//xdr:cNvPr", namespaces={"xdr": NS_XDR}
                )
            ]
            or [0]
        ) + 1
        row, column = self.worksheet._picture_anchor or (1, 1)
        anchor = etree.SubElement(
            self._drawing_root, _qname(NS_XDR, "oneCellAnchor")
        )
        origin = etree.SubElement(anchor, _qname(NS_XDR, "from"))
        for tag, value in (
            ("col", column - 1),
            ("colOff", 0),
            ("row", row - 1),
            ("rowOff", 0),
        ):
            etree.SubElement(origin, _qname(NS_XDR, tag)).text = str(value)
        etree.SubElement(
            anchor,
            _qname(NS_XDR, "ext"),
            cx=str(int(width * 9525 * scale)),
            cy=str(int(height * 9525 * scale)),
        )
        pic = etree.SubElement(anchor, _qname(NS_XDR, "pic"))
        nv_pic = etree.SubElement(pic, _qname(NS_XDR, "nvPicPr"))
        etree.SubElement(
            nv_pic,
            _qname(NS_XDR, "cNvPr"),
            id=str(picture_id),
            name=name,
        )
        etree.SubElement(nv_pic, _qname(NS_XDR, "cNvPicPr"))
        blip_fill = etree.SubElement(pic, _qname(NS_XDR, "blipFill"))
        etree.SubElement(
            blip_fill,
            _qname(NS_A, "blip"),
            {_qname(NS_REL, "embed"): rel_id},
        )
        stretch = etree.SubElement(blip_fill, _qname(NS_A, "stretch"))
        etree.SubElement(stretch, _qname(NS_A, "fillRect"))
        shape = etree.SubElement(pic, _qname(NS_XDR, "spPr"))
        transform = etree.SubElement(shape, _qname(NS_A, "xfrm"))
        etree.SubElement(transform, _qname(NS_A, "off"), x="0", y="0")
        etree.SubElement(
            transform,
            _qname(NS_A, "ext"),
            cx=str(int(width * 9525 * scale)),
            cy=str(int(height * 9525 * scale)),
        )
        geometry = etree.SubElement(
            shape, _qname(NS_A, "prstGeom"), prst="rect"
        )
        etree.SubElement(geometry, _qname(NS_A, "avLst"))
        etree.SubElement(anchor, _qname(NS_XDR, "clientData"))
        self._dirty = True

    @staticmethod
    def _image_size(path: Path) -> tuple[int, int]:
        try:
            from PIL import Image

            with Image.open(path) as image:
                return image.size
        except (ImportError, OSError):
            return 800, 450

    def _load(self, create: bool) -> None:
        if self._drawing_root is not None:
            return
        workbook = self.worksheet.workbook
        drawing = self.worksheet.root.find(_qname(NS_MAIN, "drawing"))
        sheet_rels_part = _relationship_part(self.worksheet.part)
        sheet_rels_root = workbook._relationship_root(sheet_rels_part, create)

        if drawing is not None and sheet_rels_root is not None:
            rel_id = drawing.get(_qname(NS_REL, "id"))
            relationship = sheet_rels_root.find(
                f"{_qname(NS_PKG_REL, 'Relationship')}[@Id='{rel_id}']"
            )
            if relationship is not None:
                self._drawing_part = _normalise_part(
                    self.worksheet.part, relationship.get("Target")
                )
        elif create and sheet_rels_root is not None:
            self._drawing_part = workbook._next_part(
                "xl/drawings/drawing", ".xml"
            )
            rel_id = workbook._next_relationship_id(sheet_rels_root)
            etree.SubElement(
                sheet_rels_root,
                _qname(NS_PKG_REL, "Relationship"),
                Id=rel_id,
                Type=REL_DRAWING,
                Target=f"../drawings/{posixpath.basename(self._drawing_part)}",
            )
            drawing = etree.Element(
                _qname(NS_MAIN, "drawing"),
                {_qname(NS_REL, "id"): rel_id},
            )
            ext_lst = self.worksheet.root.find(_qname(NS_MAIN, "extLst"))
            if ext_lst is None:
                self.worksheet.root.append(drawing)
            else:
                ext_lst.addprevious(drawing)
            workbook._ensure_override_content_type(
                f"/{self._drawing_part}",
                "application/vnd.openxmlformats-officedocument.drawing+xml",
            )
            workbook._new_parts.add(self._drawing_part)
            self.worksheet.dirty = True

        if self._drawing_part is None:
            return
        if self._drawing_part in workbook._parts:
            self._drawing_root = etree.fromstring(
                workbook._parts[self._drawing_part]
            )
        else:
            self._drawing_root = etree.Element(
                _qname(NS_XDR, "wsDr"),
                nsmap={"xdr": NS_XDR, "a": NS_A},
            )
        self._drawing_rels_part = _relationship_part(self._drawing_part)
        self._drawing_rels_root = workbook._relationship_root(
            self._drawing_rels_part, create=True
        )

    def flush(self) -> None:
        if self._drawing_root is None:
            return
        if self._dirty or self._drawing_part in self.worksheet.workbook._new_parts:
            assert self._drawing_part is not None
            assert self._drawing_rels_part is not None
            assert self._drawing_rels_root is not None
            self.worksheet.workbook._parts[self._drawing_part] = _serialise(
                self._drawing_root
            )
            self.worksheet.workbook._parts[
                self._drawing_rels_part
            ] = _serialise(self._drawing_rels_root)
            self.worksheet.workbook._new_parts.add(self._drawing_rels_part)


class OoxmlRange:
    """Plage de cellules exposant les opérations utilisées par le dashboard."""

    def __init__(
        self,
        worksheet: "OoxmlWorksheet",
        row: int,
        column: int,
        rows: int = 1,
        columns: int = 1,
    ) -> None:
        self.worksheet = worksheet
        self.row = row
        self.column = column
        self.rows = rows
        self.columns = columns
        self._include_index = False
        self._include_header = True

    @property
    def api(self) -> "OoxmlRange":
        return self

    @property
    def left(self) -> float:
        self.worksheet._picture_anchor = (self.row, self.column)
        return float((self.column - 1) * 64)

    @property
    def top(self) -> float:
        self.worksheet._picture_anchor = (self.row, self.column)
        return float((self.row - 1) * 20)

    @property
    def HorizontalAlignment(self) -> None:
        return None

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, alignment: Any) -> None:
        value = str(alignment).lower()
        horizontal = "center" if "center" in value else value
        self.worksheet.apply_style(
            self.row,
            self.column,
            self.rows,
            self.columns,
            horizontal=horizontal,
        )

    @property
    def number_format(self) -> None:
        return None

    @number_format.setter
    def number_format(self, code: str) -> None:
        self.worksheet.apply_style(
            self.row,
            self.column,
            self.rows,
            self.columns,
            number_format=code,
        )

    @property
    def value(self) -> Any:
        return self.worksheet.cell_value(self.row, self.column)

    @value.setter
    def value(self, value: Any) -> None:
        if isinstance(value, pd.DataFrame):
            frame = value.copy()
            if self._include_index:
                frame = frame.reset_index()
            matrix: list[list[Any]] = []
            if self._include_header:
                matrix.append(list(frame.columns))
            matrix.extend(frame.itertuples(index=False, name=None))
            self.worksheet.write_matrix(self.row, self.column, matrix)
        elif isinstance(value, pd.Series):
            self.worksheet.write_matrix(
                self.row, self.column, [[item] for item in value.tolist()]
            )
        elif (
            isinstance(value, (list, tuple))
            and value
            and isinstance(value[0], (list, tuple))
        ):
            self.worksheet.write_matrix(self.row, self.column, value)
        else:
            self.worksheet.write_cell(self.row, self.column, value)

    @property
    def formula(self) -> None:
        return None

    @formula.setter
    def formula(self, value: Any) -> None:
        if (
            isinstance(value, (list, tuple))
            and value
            and isinstance(value[0], (list, tuple))
        ):
            self.worksheet.write_formula_matrix(self.row, self.column, value)
        else:
            self.worksheet.write_formula(self.row, self.column, value)

    def options(self, index: bool = False, header: bool = True) -> "OoxmlRange":
        self._include_index = index
        self._include_header = header
        return self

    def resize(self, rows: int, columns: int) -> "OoxmlRange":
        return OoxmlRange(self.worksheet, self.row, self.column, rows, columns)

    def offset(self, rows: int, columns: int) -> "OoxmlRange":
        return OoxmlRange(
            self.worksheet, self.row + rows, self.column + columns
        )

    def expand(self) -> "OoxmlRange":
        max_row, max_column = self.worksheet.populated_bounds(
            self.row, self.column
        )
        return OoxmlRange(
            self.worksheet,
            self.row,
            self.column,
            max(1, max_row - self.row + 1),
            max(1, max_column - self.column + 1),
        )

    def get_address(self) -> str:
        return _cell_address(self.row, self.column, absolute=True)


class _UsedRange:
    def __init__(self, worksheet: "OoxmlWorksheet") -> None:
        self.worksheet = worksheet

    def clear_contents(self) -> None:
        self.worksheet.clear_contents()


class OoxmlWorksheet:
    """Feuille XLSX modifiée en mémoire."""

    def __init__(self, workbook: "OoxmlWorkbook", name: str, part: str) -> None:
        self.workbook = workbook
        self.name = name
        self.part = part
        self.root = etree.fromstring(workbook._parts[part])
        self.dirty = False
        self._picture_anchor: tuple[int, int] | None = None
        self.pictures = _Pictures(self)

    @property
    def used_range(self) -> _UsedRange:
        return _UsedRange(self)

    def range(self, *args: Any) -> OoxmlRange:
        if len(args) == 1 and isinstance(args[0], str):
            row, column = _parse_address(args[0])
        elif len(args) == 1 and isinstance(args[0], tuple):
            row, column = args[0]
        elif len(args) == 2:
            row, column = args
        else:
            raise TypeError("La plage attend une adresse A1 ou une paire ligne/colonne.")
        return OoxmlRange(self, int(row), int(column))

    def _sheet_data(self) -> etree._Element:
        sheet_data = self.root.find(_qname(NS_MAIN, "sheetData"))
        if sheet_data is None:
            sheet_data = etree.Element(_qname(NS_MAIN, "sheetData"))
            self.root.insert(0, sheet_data)
        return sheet_data

    def _row(self, row_number: int) -> etree._Element:
        sheet_data = self._sheet_data()
        row = sheet_data.find(
            f"{_qname(NS_MAIN, 'row')}[@r='{row_number}']"
        )
        if row is None:
            row = etree.Element(_qname(NS_MAIN, "row"), r=str(row_number))
            for current in sheet_data:
                if int(current.get("r", "0")) > row_number:
                    current.addprevious(row)
                    break
            else:
                sheet_data.append(row)
        return row

    def _cell(self, row_number: int, column_number: int) -> etree._Element:
        row = self._row(row_number)
        address = _cell_address(row_number, column_number)
        cell = row.find(f"{_qname(NS_MAIN, 'c')}[@r='{address}']")
        if cell is None:
            cell = etree.Element(_qname(NS_MAIN, "c"), r=address)
            for current in row:
                current_column = _parse_address(current.get("r"))[1]
                if current_column > column_number:
                    current.addprevious(cell)
                    break
            else:
                row.append(cell)
        return cell

    @staticmethod
    def _clear_cell(cell: etree._Element) -> None:
        for child in list(cell):
            if child.tag in {
                _qname(NS_MAIN, "f"),
                _qname(NS_MAIN, "v"),
                _qname(NS_MAIN, "is"),
            }:
                cell.remove(child)
        cell.attrib.pop("t", None)

    def clear_contents(self) -> None:
        for cell in self.root.iter(_qname(NS_MAIN, "c")):
            self._clear_cell(cell)
        self.dirty = True

    def write_cell(self, row: int, column: int, value: Any) -> None:
        cell = self._cell(row, column)
        self._clear_cell(cell)
        value = _python_value(value)
        if _is_blank(value):
            self.dirty = True
            return

        if isinstance(value, (pd.Timestamp, dt.datetime, dt.date)):
            timestamp = pd.Timestamp(value)
            if timestamp.tzinfo is not None:
                timestamp = timestamp.tz_convert(None)
            epoch = pd.Timestamp("1904-01-01" if self.workbook.date_1904 else "1899-12-30")
            serial = (timestamp - epoch).total_seconds() / 86400
            etree.SubElement(cell, _qname(NS_MAIN, "v")).text = f"{serial:.12g}"
            base_style = int(cell.get("s", "0"))
            cell.set(
                "s",
                str(
                    self.workbook.styles.variant(
                        base_style, number_format="yyyy-mm-dd"
                    )
                ),
            )
        elif isinstance(value, bool):
            cell.set("t", "b")
            etree.SubElement(cell, _qname(NS_MAIN, "v")).text = "1" if value else "0"
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
                self.dirty = True
                return
            etree.SubElement(cell, _qname(NS_MAIN, "v")).text = f"{value:.15g}"
        else:
            cell.set("t", "inlineStr")
            inline = etree.SubElement(cell, _qname(NS_MAIN, "is"))
            text = etree.SubElement(inline, _qname(NS_MAIN, "t"))
            string = str(value)
            if string.startswith(" ") or string.endswith(" "):
                text.set(
                    "{http://www.w3.org/XML/1998/namespace}space", "preserve"
                )
            text.text = string
        self.dirty = True

    def write_matrix(
        self, row: int, column: int, matrix: Iterable[Iterable[Any]]
    ) -> None:
        for row_offset, values in enumerate(matrix):
            for column_offset, value in enumerate(values):
                self.write_cell(row + row_offset, column + column_offset, value)

    def write_formula(self, row: int, column: int, formula: Any) -> None:
        cell = self._cell(row, column)
        self._clear_cell(cell)
        cell.attrib.pop("cm", None)
        cell.attrib.pop("vm", None)
        text = str(formula or "")
        if text.startswith("="):
            text = text[1:]
        if text:
            etree.SubElement(cell, _qname(NS_MAIN, "f")).text = text
            etree.SubElement(cell, _qname(NS_MAIN, "v"))
        self.dirty = True

    def write_formula_matrix(
        self, row: int, column: int, matrix: Iterable[Iterable[Any]]
    ) -> None:
        for row_offset, values in enumerate(matrix):
            for column_offset, formula in enumerate(values):
                self.write_formula(
                    row + row_offset,
                    column + column_offset,
                    formula,
                )

    def cell_value(self, row: int, column: int) -> Any:
        address = _cell_address(row, column)
        cell = self.root.find(
            f".//{_qname(NS_MAIN, 'c')}[@r='{address}']"
        )
        if cell is None:
            return None
        inline = cell.find(
            f"{_qname(NS_MAIN, 'is')}/{_qname(NS_MAIN, 't')}"
        )
        if inline is not None:
            return inline.text
        value = cell.find(_qname(NS_MAIN, "v"))
        return value.text if value is not None else None

    def populated_bounds(self, start_row: int, start_column: int) -> tuple[int, int]:
        max_row, max_column = start_row, start_column
        for cell in self.root.iter(_qname(NS_MAIN, "c")):
            if not list(cell):
                continue
            row, column = _parse_address(cell.get("r"))
            if row >= start_row and column >= start_column:
                max_row = max(max_row, row)
                max_column = max(max_column, column)
        return max_row, max_column

    def apply_style(
        self,
        row: int,
        column: int,
        rows: int,
        columns: int,
        number_format: str | None = None,
        horizontal: str | None = None,
    ) -> None:
        for current_row in range(row, row + rows):
            for current_column in range(column, column + columns):
                cell = self._cell(current_row, current_column)
                base_style = int(cell.get("s", "0"))
                cell.set(
                    "s",
                    str(
                        self.workbook.styles.variant(
                            base_style, number_format, horizontal
                        )
                    ),
                )
        self.dirty = True

    def autofit(self) -> None:
        widths: dict[int, int] = {}
        for cell in self.root.iter(_qname(NS_MAIN, "c")):
            if not list(cell):
                continue
            _, column = _parse_address(cell.get("r"))
            value = self.cell_value(*_parse_address(cell.get("r")))
            widths[column] = max(widths.get(column, 0), len(str(value or "")))
        if not widths:
            return
        columns_node = self.root.find(_qname(NS_MAIN, "cols"))
        if columns_node is None:
            columns_node = etree.Element(_qname(NS_MAIN, "cols"))
            sheet_data = self._sheet_data()
            sheet_data.addprevious(columns_node)
        for column, width in widths.items():
            etree.SubElement(
                columns_node,
                _qname(NS_MAIN, "col"),
                min=str(column),
                max=str(column),
                width=str(min(max(width + 2, 8), 80)),
                customWidth="1",
                bestFit="1",
            )
        self.dirty = True

    def hide_columns(self, start_column: int, end_column: int) -> None:
        columns_node = self.root.find(_qname(NS_MAIN, "cols"))
        if columns_node is None:
            columns_node = etree.Element(_qname(NS_MAIN, "cols"))
            self._sheet_data().addprevious(columns_node)
        etree.SubElement(
            columns_node,
            _qname(NS_MAIN, "col"),
            min=str(start_column),
            max=str(end_column),
            hidden="1",
        )
        self.dirty = True

    def flush(self) -> None:
        self.pictures.flush()
        if self.dirty:
            dimension = self.root.find(_qname(NS_MAIN, "dimension"))
            if dimension is not None:
                max_row, max_column = self.populated_bounds(1, 1)
                dimension.set(
                    "ref",
                    f"A1:{_cell_address(max_row, max_column)}",
                )
            self.workbook._parts[self.part] = _serialise(self.root)


class _SheetCollection:
    def __init__(self, worksheets: list[OoxmlWorksheet]) -> None:
        self._worksheets = worksheets
        self._by_name = {worksheet.name: worksheet for worksheet in worksheets}

    def __iter__(self) -> Iterable[OoxmlWorksheet]:
        return iter(self._worksheets)

    def __getitem__(self, name: str) -> OoxmlWorksheet:
        return self._by_name[name]


class OoxmlWorkbook:
    """Classeur XLSX fondé sur un modèle existant."""

    def __init__(self, template_path: str | os.PathLike[str]) -> None:
        self.template_path = Path(template_path)
        self._parts: dict[str, bytes] = {}
        self._infos: dict[str, zipfile.ZipInfo] = {}
        self._new_parts: set[str] = set()
        self._relationship_roots: dict[str, etree._Element] = {}
        with zipfile.ZipFile(self.template_path, "r") as archive:
            for info in archive.infolist():
                self._infos[info.filename] = info
                self._parts[info.filename] = archive.read(info.filename)

        self._workbook_root = etree.fromstring(self._parts["xl/workbook.xml"])
        self._workbook_rels_root = etree.fromstring(
            self._parts["xl/_rels/workbook.xml.rels"]
        )
        workbook_properties = self._workbook_root.find(
            _qname(NS_MAIN, "workbookPr")
        )
        self.date_1904 = (
            workbook_properties is not None
            and workbook_properties.get("date1904") in {"1", "true", "True"}
        )
        self.styles = _StyleManager(self)
        self.sheets = _SheetCollection(self._load_sheets())

    def _load_sheets(self) -> list[OoxmlWorksheet]:
        relationships = {
            relationship.get("Id"): relationship.get("Target")
            for relationship in self._workbook_rels_root
        }
        worksheets = []
        sheets_node = self._workbook_root.find(_qname(NS_MAIN, "sheets"))
        if sheets_node is None:
            return worksheets
        for sheet in sheets_node:
            rel_id = sheet.get(_qname(NS_REL, "id"))
            part = _normalise_part("xl/workbook.xml", relationships[rel_id])
            worksheets.append(OoxmlWorksheet(self, sheet.get("name"), part))
        return worksheets

    def _relationship_root(
        self, part: str, create: bool
    ) -> etree._Element | None:
        if part in self._relationship_roots:
            return self._relationship_roots[part]
        if part in self._parts:
            root = etree.fromstring(self._parts[part])
            self._relationship_roots[part] = root
            return root
        if not create:
            return None
        root = etree.Element(
            _qname(NS_PKG_REL, "Relationships"),
            nsmap={None: NS_PKG_REL},
        )
        self._parts[part] = _serialise(root)
        self._new_parts.add(part)
        self._relationship_roots[part] = root
        return root

    @staticmethod
    def _next_relationship_id(root: etree._Element) -> str:
        used = {
            int(match.group(1))
            for item in root
            if (match := re.match(r"rId(\d+)$", item.get("Id", "")))
        }
        number = 1
        while number in used:
            number += 1
        return f"rId{number}"

    def _next_part(self, prefix: str, suffix: str) -> str:
        number = 1
        while f"{prefix}{number}{suffix}" in self._parts:
            number += 1
        return f"{prefix}{number}{suffix}"

    def _content_types_root(self) -> etree._Element:
        return etree.fromstring(self._parts["[Content_Types].xml"])

    def _ensure_default_content_type(
        self, extension: str, content_type: str
    ) -> None:
        root = self._content_types_root()
        found = root.find(
            f"{_qname(NS_CT, 'Default')}[@Extension='{extension}']"
        )
        if found is None:
            etree.SubElement(
                root,
                _qname(NS_CT, "Default"),
                Extension=extension,
                ContentType=content_type,
            )
            self._parts["[Content_Types].xml"] = _serialise(root)

    def _ensure_override_content_type(
        self, part_name: str, content_type: str
    ) -> None:
        root = self._content_types_root()
        found = root.find(
            f"{_qname(NS_CT, 'Override')}[@PartName='{part_name}']"
        )
        if found is None:
            etree.SubElement(
                root,
                _qname(NS_CT, "Override"),
                PartName=part_name,
                ContentType=content_type,
            )
            self._parts["[Content_Types].xml"] = _serialise(root)

    def _remove_calc_chain(self) -> None:
        self._parts.pop("xl/calcChain.xml", None)
        for relationship in list(self._workbook_rels_root):
            if relationship.get("Type") == REL_CALC_CHAIN:
                self._workbook_rels_root.remove(relationship)
        content_types = self._content_types_root()
        for override in list(content_types):
            if override.get("PartName") == "/xl/calcChain.xml":
                content_types.remove(override)
        self._parts["[Content_Types].xml"] = _serialise(content_types)

    def _prepare_calculation(self) -> None:
        calc_properties = self._workbook_root.find(_qname(NS_MAIN, "calcPr"))
        if calc_properties is None:
            calc_properties = etree.SubElement(
                self._workbook_root, _qname(NS_MAIN, "calcPr")
            )
        calc_properties.set("calcMode", "auto")
        calc_properties.set("fullCalcOnLoad", "1")
        calc_properties.set("forceFullCalc", "1")
        self._remove_calc_chain()

    def save(self, output_path: str | os.PathLike[str]) -> None:
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.resolve() == self.template_path.resolve():
            raise ValueError("Le modèle et la sortie doivent être deux fichiers distincts.")

        for worksheet in self.sheets:
            worksheet.flush()
        self.styles.flush()
        self._prepare_calculation()
        self._parts["xl/workbook.xml"] = _serialise(self._workbook_root)
        self._parts["xl/_rels/workbook.xml.rels"] = _serialise(
            self._workbook_rels_root
        )
        for part, root in self._relationship_roots.items():
            self._parts[part] = _serialise(root)

        descriptor, temporary_name = tempfile.mkstemp(
            suffix=".xlsx", dir=output.parent
        )
        os.close(descriptor)
        try:
            with zipfile.ZipFile(temporary_name, "w") as archive:
                for part, data in self._parts.items():
                    info = self._infos.get(part)
                    if info is None:
                        archive.writestr(part, data, compress_type=zipfile.ZIP_DEFLATED)
                    else:
                        archive.writestr(info, data)
            os.replace(temporary_name, output)
        finally:
            if os.path.exists(temporary_name):
                os.unlink(temporary_name)
