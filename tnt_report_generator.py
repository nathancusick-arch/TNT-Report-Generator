from __future__ import annotations

import collections
import copy
import csv
import datetime as dt
import hashlib
import html
import io
import json
import math
import posixpath
import re
import sys
import threading
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from functools import lru_cache
from typing import Any
from zoneinfo import ZoneInfo
from xml.sax.saxutils import escape, quoteattr

import pandas as pd
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import from_excel, to_excel

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
Q = lambda tag: "{" + NS + "}" + tag
INPUTS = ("Weekly Input", "Cumulative Input", "Amazon Data")
SUMMARIES = ("Weekly Summary Table", "Cumulative Summary ", "Amazon Summary(Excluded)")
CLIENT_SHEETS = ("Amazon Summary(Excluded)", "Weekly Summary Table", "Cumulative Summary ", "Headline Figures")
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class ReportError(ValueError):
    """A validation problem that can be displayed directly to the operator."""


def xml_fragment(element: ET.Element) -> str:
    """Write worksheet fragments independently of ET's global prefix registry.

    Streamlit and other libraries may register 's' instead of 'ns0'. Main-sheet
    element names inherit the surrounding worksheet namespace; qualified
    attributes retain their automatically declared namespaces.
    """
    element = copy.deepcopy(element)
    for node in element.iter():
        if isinstance(node.tag, str) and node.tag.startswith("{" + NS + "}"):
            node.tag = node.tag[len(NS) + 2:]
    return ET.tostring(element, encoding="unicode")


def validate_xlsx(content: bytes):
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            for name in archive.namelist():
                if name.endswith((".xml", ".rels")):
                    with archive.open(name) as part:
                        for _, element in ET.iterparse(part, events=("end",)):
                            element.clear()
    except (ET.ParseError, zipfile.BadZipFile) as error:
        raise ReportError(f"The generated workbook failed its integrity check: {error}") from error


def clean(value: Any) -> str:
    return "" if value is None or value is pd.NA or (isinstance(value, float) and math.isnan(value)) else str(value).strip()


def title_key(value: str) -> str:
    value = re.sub(r"^Q\d+\s*-\s*", "", value.strip(), flags=re.I)
    return re.sub(r"\s+", " ", value).strip(" :").casefold()


def read_csv_bytes(content: bytes) -> list[dict[str, str]]:
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            text = content.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ReportError("The CSV encoding could not be read. Save it as UTF-8 and upload it again.")
    if not text.strip():
        raise ReportError("The uploaded CSV is empty.")
    first = text.splitlines()[0]
    delimiter = ";" if first.count(";") > first.count(",") else "\t" if first.count("\t") > first.count(",") else ","
    reader = csv.DictReader(io.StringIO(text, newline=""), delimiter=delimiter)
    headers = [clean(h) for h in (reader.fieldnames or [])]
    if len(headers) != len(set(headers)):
        raise ReportError("The CSV contains duplicate column headers. Use a fresh system export.")
    result = []
    for row in reader:
        if None in row and row[None]:
            raise ReportError("A CSV row has more values than its headers. Check the export format.")
        if any(clean(v) for v in row.values()):
            result.append({clean(k): "" if v is None else v for k, v in row.items() if k is not None})
    return result


def parse_date(value: Any) -> dt.date:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)):
        return from_excel(value).date()
    text = clean(value)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return pd.to_datetime(text, dayfirst=not bool(re.match(r"^\d{4}-", text))).date()
    except (ValueError, TypeError):
        raise ReportError(f"Could not read visit date: {text or '(blank)' }.") from None


def excel_date(value: Any) -> float:
    return to_excel(dt.datetime.combine(parse_date(value), dt.time()))


def number_or_text(value: str) -> Any:
    if not clean(value):
        return None
    try:
        value_number = float(value)
        return int(value_number) if value_number.is_integer() else value_number
    except ValueError:
        return value


def addr_parts(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", address)
    if not match:
        raise ReportError(f"Invalid cell reference: {address}")
    return int(match[2]), column_index_from_string(match[1])


@dataclass(slots=True)
class Cell:
    value: Any = None
    formula: str | None = None
    style: str | None = None
    error: bool = False
    node: Any = None
    calculated: bool = False
    trusted_cache: bool = False
    rich: int | None = None


class Sheet:
    def __init__(self, book: "Book", name: str, path: str):
        self.book, self.name, self.path = book, name, path
        original = book.archive.read(path)
        outer = re.sub(rb"<sheetData(?:\s[^>]*)?>.*?</sheetData>", b"__TNT_SHEET_DATA__", original, count=1, flags=re.S)
        self.original = outer.decode("utf-8")
        self.root = ET.fromstring(outer.replace(b"__TNT_SHEET_DATA__", b"<sheetData/>"))
        self.rows = {}
        self.row_columns = {}
        self.empty_styles = {}
        self.styled_max_col = 0
        self.cells: dict[tuple[int, int], Cell] = {}
        self.shared_formulas = {}
        self.max_row = self.max_col = 0
        self.dirty = False
        self.filter_ref = None
        self.extra_columns: list[tuple[int, int]] = []
        for _, row in ET.iterparse(io.BytesIO(original), events=("end",)):
            if row.tag != Q("row"):
                continue
            row_number = int(row.get("r"))
            self.rows[row_number] = dict(row.attrib)
            self.row_columns[row_number] = []
            for node in row:
                address = node.attrib["r"]
                r, c = addr_parts(address)
                self.row_columns[r].append(c)
                self.styled_max_col = max(self.styled_max_col, c)
                kind = node.get("t", "n")
                v = node.find(Q("v"))
                raw = v.text if v is not None else None
                if kind == "s":
                    value = book.strings[int(raw)] if raw is not None else None
                elif kind == "inlineStr":
                    value = "".join(t.text or "" for t in node.findall(".//" + Q("t")))
                elif kind in ("str", "e", "d"):
                    value = raw
                elif kind == "b":
                    value = raw == "1"
                else:
                    value = number_or_text(raw) if raw is not None else None
                f = node.find(Q("f"))
                formula = None
                if f is not None:
                    if f.text:
                        formula = "=" + f.text
                        if f.get("t") == "shared":
                            self.shared_formulas[f.get("si")] = (address, formula)
                    elif f.get("t") == "shared":
                        origin, base = self.shared_formulas[f.get("si")]
                        formula = Translator(base, origin=origin).translate_formula(address)
                style = sys.intern(node.get("s")) if node.get("s") else None
                attrs = {k: v for k, v in node.attrib.items() if k not in ("r", "s", "t")} or None
                if value is not None or formula or r <= 6 or attrs:
                    rich = int(raw) if kind == "s" and raw is not None and int(raw) in book.rich_strings else None
                    cached = not formula or (v is not None and (raw is not None or kind == "str"))
                    self.cells[(r, c)] = Cell(value, formula, style, kind == "e", attrs, trusted_cache=cached, rich=rich)
                else:
                    self.empty_styles[(r, c)] = style
                if value is not None or formula:
                    self.max_row, self.max_col = max(self.max_row, r), max(self.max_col, c)
            row.clear()

    def get(self, row: int, col: int) -> Any:
        cell = self.cells.get((row, col))
        return cell.value if cell else None

    def audit_rows(self) -> list[int]:
        return [r for r in range(3, self.max_row + 1) if clean(self.get(r, 4))]

    def set(self, row: int, col: int, value: Any = None, formula: str | None = None, template_row: int | None = None, template_col: int | None = None, error: bool = False):
        old = self.cells.get((row, col))
        if old is None:
            template = self.cells.get((template_row or row, template_col or col))
            style = self.empty_styles.get((row, col))
            if (row, col) not in self.empty_styles:
                style = template.style if template else self.empty_styles.get((template_row or row, template_col or col))
            old = Cell(style=style)
            if row not in self.rows:
                original_row = self.rows.get(template_row or row)
                attrs = {k: v for k, v in (original_row or {}).items() if k not in ("r", "spans", "hidden")}
                attrs["r"] = str(row)
                self.rows[row] = attrs
                self.row_columns[row] = []
            if col not in self.row_columns[row]:
                self.row_columns[row].append(col)
            self.cells[(row, col)] = old
        old.value, old.formula, old.error, old.calculated = value, formula, error, False
        old.trusted_cache = False
        old.rich = None
        self.max_row, self.max_col = max(self.max_row, row), max(self.max_col, col)
        self.dirty = True

    def clear_from(self, first_row: int):
        for (r, c), cell in list(self.cells.items()):
            if r >= first_row and (cell.value is not None or cell.formula):
                self.set(r, c)

    @staticmethod
    def _write_node(cell: Cell, value: Any, formula: str | None):
        node = cell.node
        for child in list(node):
            if child.tag in (Q("v"), Q("f"), Q("is")):
                node.remove(child)
        node.attrib.pop("t", None)
        if formula:
            ET.SubElement(node, Q("f")).text = formula.lstrip("=")
        if isinstance(value, ExcelFault):
            node.set("t", "e")
            ET.SubElement(node, Q("v")).text = value.code
        elif formula or value is not None:
            if isinstance(value, bool):
                node.set("t", "b")
                ET.SubElement(node, Q("v")).text = "1" if value else "0"
            elif isinstance(value, (float, int)) and not isinstance(value, bool):
                if not math.isfinite(value):
                    raise ReportError("A report calculation returned a non-finite number.")
                ET.SubElement(node, Q("v")).text = str(value)
            elif formula:
                node.set("t", "str")
                ET.SubElement(node, Q("v")).text = "" if value is None else str(value)
            else:
                node.set("t", "inlineStr")
                t = ET.SubElement(ET.SubElement(node, Q("is")), Q("t"))
                t.text = str(value)
                if str(value) != str(value).strip():
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def xml(self, values_only: bool = False, end_row: int | None = None) -> bytes:
        body = io.BytesIO(); body.write(b"<sheetData>")
        for row_number in sorted(self.rows):
            if end_row is not None and row_number > end_row:
                continue
            opening = xml_fragment(ET.Element(Q("row"), self.rows[row_number]))
            parts = [opening[:-2] + ">"]
            for col in sorted(self.row_columns[row_number]):
                source = self.cells.get((row_number, col))
                attrs = dict(source.node or {}) if source else {}
                attrs["r"] = f"{get_column_letter(col)}{row_number}"
                style = source.style if source else self.empty_styles.get((row_number, col))
                if style:
                    attrs["s"] = style
                value = source.value if source else None
                if source and source.error and not isinstance(value, ExcelFault):
                    value = ExcelFault(str(value))
                formula = None if values_only or source is None else source.formula
                if any(key.startswith("{") for key in attrs):
                    node = ET.Element(Q("c"), attrs)
                    self._write_node(Cell(node=node), value, formula)
                    parts.append(xml_fragment(node))
                    continue
                prefix = "<c " + " ".join(key + "=" + quoteattr(str(val)) for key, val in attrs.items())
                f = "<f>" + escape(formula.lstrip("=")) + "</f>" if formula else ""
                if source and source.rich is not None and not formula:
                    if values_only:
                        parts.append(prefix + ' t="inlineStr">' + self.book.rich_strings[source.rich] + "</c>")
                    else:
                        parts.append(prefix + ' t="s"><v>' + str(source.rich) + "</v></c>")
                elif isinstance(value, ExcelFault):
                    parts.append(prefix + ' t="e">' + f + "<v>" + escape(value.code) + "</v></c>")
                elif isinstance(value, bool):
                    parts.append(prefix + ' t="b">' + f + ("<v>1</v></c>" if value else "<v>0</v></c>"))
                elif isinstance(value, (int, float)):
                    if not math.isfinite(value):
                        raise ReportError("A report calculation returned a non-finite number.")
                    parts.append(prefix + ">" + f + "<v>" + str(value) + "</v></c>")
                elif formula:
                    parts.append(prefix + ' t="str">' + f + "<v>" + escape("" if value is None else str(value)) + "</v></c>")
                elif value is not None:
                    text = str(value)
                    space = ' xml:space="preserve"' if text != text.strip() else ""
                    parts.append(prefix + ' t="inlineStr"><is><t' + space + ">" + escape(text) + "</t></is></c>")
                else:
                    parts.append(prefix + "/>")
            parts.append("</row>")
            body.write("".join(parts).encode())
        body.write(b"</sheetData>")
        xml = self.original
        maximum_row = end_row if end_row is not None else max(self.max_row, max(self.rows, default=1))
        dimension = f"A1:{get_column_letter(max(self.max_col, self.styled_max_col))}{max(maximum_row, 1)}"
        xml = re.sub(r'<dimension\b[^>]*/>', f'<dimension ref="{dimension}"/>', xml, count=1)
        if self.filter_ref:
            xml = re.sub(r'(<autoFilter\b[^>]*\bref=")[^"]+(\")', lambda m: m[1] + self.filter_ref + m[2], xml, count=1)
        if self.extra_columns:
            columns = self.root.find(Q("cols"))
            if columns is not None:
                columns = copy.deepcopy(columns)
                for target, source in self.extra_columns:
                    existing = next((c for c in columns if int(c.get("min")) <= target <= int(c.get("max"))), None)
                    if existing is not None:
                        continue
                    source_col = next((c for c in columns if int(c.get("min")) <= source <= int(c.get("max"))), None)
                    if source_col is not None:
                        new = copy.deepcopy(source_col)
                        new.set("min", str(target)); new.set("max", str(target)); columns.append(new)
                text = xml_fragment(columns)
                xml = re.sub(r"<(?:[\w.-]+:)?cols\b[^>]*>.*?</(?:[\w.-]+:)?cols>", lambda _: text, xml, count=1, flags=re.S)
        if values_only:
            xml = re.sub(r"<legacyDrawing\b[^>]*/>", "", xml)
        prefix, suffix = xml.split("__TNT_SHEET_DATA__", 1)
        return prefix.encode() + body.getvalue() + suffix.encode()


class Book:
    def __init__(self, content: bytes):
        try:
            self.archive = zipfile.ZipFile(io.BytesIO(content))
            self.workbook_xml = self.archive.read("xl/workbook.xml").decode()
        except (zipfile.BadZipFile, KeyError):
            raise ReportError("Upload an Excel .xlsx LIVE report.") from None
        workbook = ET.fromstring(self.workbook_xml)
        if workbook.find(Q("workbookPr")) is not None and workbook.find(Q("workbookPr")).get("date1904") in ("1", "true"):
            raise ReportError("This workbook uses the 1904 date system. Use the supplied TNT LIVE template.")
        self.paths = {}
        self.states = {}
        rels = ET.fromstring(self.archive.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels}
        for sheet in workbook.find(Q("sheets")):
            target = targets[sheet.get("{" + REL + "}id")]
            self.paths[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            self.states[sheet.get("name")] = sheet.get("state", "visible")
        self.aliases = {name.strip(): name for name in self.paths}
        self.strings = []
        self.rich_strings = {}
        if "xl/sharedStrings.xml" in self.archive.namelist():
            strings = ET.fromstring(self.archive.read("xl/sharedStrings.xml"))
            self.strings = ["".join(t.text or "" for t in item.iter(Q("t"))) for item in strings]
            for index, item in enumerate(strings):
                if item.find(Q("r")) is not None:
                    inline = ET.Element(Q("is"))
                    for child in item:
                        inline.append(copy.deepcopy(child))
                    self.rich_strings[index] = xml_fragment(inline)
        self.loaded: dict[str, Sheet] = {}

    def sheet(self, name: str) -> Sheet:
        actual = self.aliases.get(name.strip())
        if actual is None:
            raise ReportError(f"The LIVE report is missing the '{name.strip()}' sheet.")
        if actual not in self.loaded:
            self.loaded[actual] = Sheet(self, actual, self.paths[actual])
        return self.loaded[actual]

    def live_bytes(self) -> bytes:
        changes = {sheet.path: sheet.xml() for sheet in self.loaded.values() if sheet.dirty}
        workbook = re.sub(r"<calcPr\b[^>]*/>", '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>', self.workbook_xml)
        ordered_names = list(self.paths)
        def filter_name(match):
            index = int(match[2])
            name = ordered_names[index] if index < len(ordered_names) else None
            sheet = self.loaded.get(name)
            if not sheet or not sheet.filter_ref:
                return match[0]
            absolute = re.sub(r"([A-Z]+)(\d+)", lambda m: "$" + m[1] + "$" + m[2], sheet.filter_ref)
            value = "'" + sheet.name.replace("'", "''") + "'!" + absolute
            return match[1] + escape(value) + match[4]
        workbook = re.sub(r'(<definedName\b(?=[^>]*\bname="_xlnm\._FilterDatabase")(?=[^>]*\blocalSheetId="(\d+)")[^>]*>)(.*?)(</definedName>)', filter_name, workbook, flags=re.S)
        changes["xl/workbook.xml"] = workbook.encode()
        for name, pattern in (("xl/_rels/workbook.xml.rels", r'<Relationship\b[^>]*Type="[^"]*/calcChain"[^>]*/>'), ("[Content_Types].xml", r'<Override\b[^>]*PartName="/xl/calcChain.xml"[^>]*/>')):
            changes[name] = re.sub(pattern, "", self.archive.read(name).decode()).encode()
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            for info in self.archive.infolist():
                if info.filename != "xl/calcChain.xml":
                    archive.writestr(copy.copy(info), changes[info.filename] if info.filename in changes else self.archive.read(info.filename))
        return output.getvalue()


class ExcelFault(Exception):
    def __init__(self, code: str):
        self.code = code
        super().__init__(code)


TOKEN = re.compile(r'''\s*(?:("(?:[^"]|"")*")|((?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_ ]*)!)?(?:\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+:\$?[A-Z]+))|(\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?)|([A-Za-z_][A-Za-z0-9_.]*)|(<>|<=|>=|[=<>+*/&(),%\-]))''')


@lru_cache(maxsize=2048)
def formula_ast(formula: str):
    text = formula.lstrip("=")
    tokens = []
    position = 0
    while position < len(text):
        if not text[position:].strip():
            break
        match = TOKEN.match(text, position)
        if not match:
            raise ReportError(f"The LIVE report contains an unsupported formula: {formula[:120]}")
        position = match.end()
        kind = next(i for i, value in enumerate(match.groups()) if value is not None)
        tokens.append((kind, match.group(kind + 1)))
    cursor = 0
    priority = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4}

    def expression(minimum=0):
        nonlocal cursor
        if cursor >= len(tokens):
            raise ReportError(f"Incomplete report formula: {formula[:120]}")
        kind, value = tokens[cursor]; cursor += 1
        if value in ("+", "-") and kind == 4:
            node = ("un", value, expression(5))
        elif value == "(" and kind == 4:
            node = expression()
            if cursor >= len(tokens) or tokens[cursor][1] != ")":
                raise ReportError("A report formula has an unmatched bracket.")
            cursor += 1
        elif kind == 0:
            node = ("lit", value[1:-1].replace('""', '"'))
        elif kind == 1:
            node = ("ref", value)
        elif kind == 2:
            node = ("lit", float(value))
        elif kind == 3:
            if value.upper() in ("TRUE", "FALSE"):
                node = ("lit", value.upper() == "TRUE")
            else:
                if cursor >= len(tokens) or tokens[cursor][1] != "(":
                    raise ReportError(f"Unsupported named value in report formula: {value}")
                cursor += 1
                args = []
                if tokens[cursor][1] != ")":
                    while True:
                        args.append(expression())
                        if tokens[cursor][1] != ",":
                            break
                        cursor += 1
                if tokens[cursor][1] != ")":
                    raise ReportError("A report formula has invalid arguments.")
                cursor += 1
                node = ("call", value.upper().removeprefix("_XLFN."), tuple(args))
        else:
            raise ReportError(f"Unsupported report formula token: {value}")
        while cursor < len(tokens):
            op = tokens[cursor][1]
            if op == "%":
                cursor += 1; node = ("bin", "/", node, ("lit", 100)); continue
            if op not in priority or priority[op] < minimum:
                break
            cursor += 1
            node = ("bin", op, node, expression(priority[op] + 1))
        return node

    result = expression()
    if cursor != len(tokens):
        raise ReportError(f"Could not read the complete report formula: {formula[:120]}")
    return result


@dataclass(frozen=True, slots=True)
class RangeRef:
    sheet: str
    r1: int
    c1: int
    r2: int | None
    c2: int


def excel_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return format(value, ".15g")
    return str(value)


def value_key(value: Any):
    return "" if value is None else value.casefold() if isinstance(value, str) else value


@lru_cache(maxsize=2048)
def wildcard_pattern(pattern: str):
    result, escaped = "", False
    for ch in pattern:
        if escaped:
            result += re.escape(ch); escaped = False
        elif ch == "~":
            escaped = True
        else:
            result += ".*" if ch == "*" else "." if ch == "?" else re.escape(ch)
    if escaped:
        result += re.escape("~")
    return re.compile("^" + result + "$", re.I | re.S)


def compare(left: Any, op: str, right: Any) -> bool:
    if isinstance(left, ExcelFault):
        raise left
    if isinstance(right, ExcelFault):
        raise right
    if isinstance(left, str) and isinstance(right, str):
        left, right = left.casefold(), right.casefold()
    if left is None:
        left = "" if isinstance(right, str) else 0
    if right is None:
        right = "" if isinstance(left, str) else 0
    try:
        return {"=": lambda: left == right, "<>": lambda: left != right, ">": lambda: left > right, "<": lambda: left < right, ">=": lambda: left >= right, "<=": lambda: left <= right}[op]()
    except TypeError:
        return op == "<>"


def criterion(value: Any):
    text = excel_text(value)
    match = re.match(r"^(<>|>=|<=|>|<|=)(.*)$", text, re.S)
    op, target = (match[1], match[2]) if match else ("=", value)
    if isinstance(target, str) and op not in ("=", "<>"):
        try:
            target = float(target)
        except ValueError:
            pass
    wildcard = isinstance(target, str) and ("*" in target or "?" in target)
    def test(actual):
        if wildcard:
            matched = bool(wildcard_pattern(target).fullmatch(excel_text(actual)))
            return not matched if op == "<>" else matched
        return compare(actual, op, target)
    return op, target, wildcard, test


class Calculator:
    """Evaluate the template's supported formulas and save their Excel caches.

    COUNTIFS uses indexed equality candidates, making venue/date ranks cheap.
    Formula text is carried from the template, not replaced by a repair script.
    """
    def __init__(self, book: Book):
        self.book = book
        self.active = set()
        self.indexes = {}
        self.lookup_tables = {}
        self.simple_if = re.compile(r'^=IF\(\s*(\'(?:[^\']|\'\')+\'!\$?[A-Z]+\$?\d+)\s*=\s*""\s*,\s*""\s*,\s*\1\s*\)$', re.I)
        self.simple_ref = re.compile(r"^=('(?:[^']|'')+'!\$?[A-Z]+\$?\d+|\$?[A-Z]+\$?\d+)$")

    def cell(self, sheet_name: str, row: int, col: int):
        sheet = self.book.sheet(sheet_name)
        source = sheet.cells.get((row, col))
        if source is None:
            return None
        if not source.formula:
            if source.error:
                raise ExcelFault(str(source.value))
            return source.value
        if source.calculated:
            if source.error:
                raise source.value if isinstance(source.value, ExcelFault) else ExcelFault(str(source.value))
            return source.value
        key = (sheet.name, row, col)
        if key in self.active:
            raise ExcelFault("#REF!")
        self.active.add(key)
        try:
            simple = self.simple_if.fullmatch(source.formula)
            direct = self.simple_ref.fullmatch(source.formula)
            if simple:
                ref = self.reference(simple[1], sheet.name)
                value = self.cell(ref.sheet, ref.r1, ref.c1)
                value = "" if value is None or value == "" else value
            elif direct:
                ref = self.reference(direct[1], sheet.name)
                value = self.cell(ref.sheet, ref.r1, ref.c1)
            else:
                value = self.scalar(self.eval(formula_ast(source.formula), sheet.name))
            value = 0 if value is None else value
            source.value, source.error, source.calculated = value, False, True
            return value
        except ExcelFault as error:
            source.value, source.error, source.calculated = error, True, True
            raise
        finally:
            self.active.remove(key)

    def reference(self, token: str, current_sheet: str) -> RangeRef:
        if "!" in token:
            sheet, token = token.rsplit("!", 1)
            sheet = sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet
        else:
            sheet = current_sheet
        sheet = self.book.sheet(sheet).name
        token = token.replace("$", "")
        if re.fullmatch(r"[A-Z]+:[A-Z]+", token):
            a, b = token.split(":")
            return RangeRef(sheet, 1, column_index_from_string(a), None, column_index_from_string(b))
        a, _, b = token.partition(":")
        r1, c1 = addr_parts(a); r2, c2 = addr_parts(b or a)
        return RangeRef(sheet, r1, c1, r2, c2)

    def bounds(self, ref: RangeRef):
        return ref.r1, min(ref.r2 or self.book.sheet(ref.sheet).max_row, self.book.sheet(ref.sheet).max_row)

    def values(self, ref: RangeRef):
        start, end = self.bounds(ref)
        return [self.cell(ref.sheet, row, col) for row in range(start, end + 1) for col in range(ref.c1, ref.c2 + 1)]

    def scalar(self, value):
        return self.cell(value.sheet, value.r1, value.c1) if isinstance(value, RangeRef) else value

    def numeric(self, value):
        value = self.scalar(value)
        if value is None or value == "":
            return 0
        try:
            return float(value)
        except (ValueError, TypeError):
            raise ExcelFault("#VALUE!") from None

    def binary(self, op, left, right):
        if isinstance(left, RangeRef) and (left.r2 is None or left.r2 != left.r1 or left.c2 != left.c1):
            left = self.values(left)
        else:
            left = self.scalar(left)
        if isinstance(right, RangeRef) and (right.r2 is None or right.r2 != right.r1 or right.c2 != right.c1):
            right = self.values(right)
        else:
            right = self.scalar(right)
        if isinstance(left, list) or isinstance(right, list):
            length = len(left) if isinstance(left, list) else len(right)
            if isinstance(left, list) and isinstance(right, list) and len(left) != len(right):
                raise ExcelFault("#VALUE!")
            return [self.binary(op, left[i] if isinstance(left, list) else left, right[i] if isinstance(right, list) else right) for i in range(length)]
        if op in ("=", "<>", ">", "<", ">=", "<="):
            return compare(left, op, right)
        if op == "&":
            return excel_text(left) + excel_text(right)
        a, b = self.numeric(left), self.numeric(right)
        if op == "/" and b == 0:
            raise ExcelFault("#DIV/0!")
        return {"+": lambda: a + b, "-": lambda: a - b, "*": lambda: a * b, "/": lambda: a / b}[op]()

    def index(self, ref: RangeRef):
        if ref not in self.indexes:
            index = collections.defaultdict(list)
            for offset, value in enumerate(self.values(ref)):
                index[value_key(value)].append(offset)
            self.indexes[ref] = index
        return self.indexes[ref]

    def countifs(self, pairs):
        ranges, tests, candidates = [], [], None
        for ref, value in pairs:
            if not isinstance(ref, RangeRef) or ref.c1 != ref.c2:
                raise ExcelFault("#VALUE!")
            op, target, wildcard, test = criterion(self.scalar(value))
            ranges.append(ref); tests.append(test)
            if op == "=" and not wildcard:
                positions = self.index(ref).get(value_key(target), [])
                if candidates is None or len(positions) < len(candidates):
                    candidates = positions
        lengths = [max(0, self.bounds(ref)[1] - ref.r1 + 1) for ref in ranges]
        if len(set(lengths)) > 1:
            raise ExcelFault("#VALUE!")
        if candidates is None:
            candidates = range(lengths[0])
        return sum(all(test(self.cell(ref.sheet, ref.r1 + offset, ref.c1)) for ref, test in zip(ranges, tests)) for offset in candidates)

    def eval(self, node, current_sheet):
        kind = node[0]
        if kind == "lit":
            return node[1]
        if kind == "ref":
            return self.reference(node[1], current_sheet)
        if kind == "un":
            value = self.eval(node[2], current_sheet)
            return self.binary("*", value, -1 if node[1] == "-" else 1)
        if kind == "bin":
            return self.binary(node[1], self.eval(node[2], current_sheet), self.eval(node[3], current_sheet))
        name, args = node[1], node[2]
        ev = lambda n: self.eval(n, current_sheet)
        if name == "IF":
            return ev(args[1]) if bool(self.scalar(ev(args[0]))) else ev(args[2]) if len(args) > 2 else False
        if name == "IFERROR":
            try:
                return self.scalar(ev(args[0]))
            except ExcelFault:
                return ev(args[1])
        values = [ev(arg) for arg in args]
        if name in ("COUNTIF", "COUNTIFS"):
            return self.countifs(list(zip(values[::2], values[1::2])))
        if name == "COUNTA":
            count = 0
            for arg in values:
                if isinstance(arg, RangeRef):
                    for row in range(arg.r1, self.bounds(arg)[1] + 1):
                        for col in range(arg.c1, arg.c2 + 1):
                            try:
                                count += self.cell(arg.sheet, row, col) is not None
                            except ExcelFault:
                                count += 1
                else:
                    count += self.scalar(arg) is not None
            return count
        if name in ("SUM", "SUMPRODUCT"):
            arrays = [self.values(v) if isinstance(v, RangeRef) else v if isinstance(v, list) else [v] for v in values]
            if name == "SUMPRODUCT" and len(arrays) > 1:
                return sum(math.prod(float(v or 0) for v in row) for row in zip(*arrays))
            return sum(float(v) for array in arrays for v in array if isinstance(v, (int, float)))
        if name == "LEFT":
            return excel_text(self.scalar(values[0]))[:int(self.numeric(values[1])) if len(values) > 1 else 1]
        if name == "VLOOKUP":
            needle, ref, offset = self.scalar(values[0]), values[1], int(self.numeric(values[2]))
            if len(values) < 4 or bool(self.scalar(values[3])):
                raise ReportError("This report uses an approximate VLOOKUP. Use the TNT LIVE template with exact lookups.")
            key = (ref.sheet, ref.r1, ref.c1, ref.r2, ref.c2)
            if key not in self.lookup_tables:
                result = {}
                for row in range(ref.r1, self.bounds(ref)[1] + 1):
                    value = self.cell(ref.sheet, row, ref.c1)
                    result.setdefault(value_key(value), row)
                self.lookup_tables[key] = result
            row = self.lookup_tables[key].get(value_key(needle))
            if row is None:
                raise ExcelFault("#N/A")
            value = self.cell(ref.sheet, row, ref.c1 + offset - 1)
            return 0 if value is None else value
        if name == "TEXT":
            value, fmt = self.scalar(values[0]), clean(self.scalar(values[1])).lower()
            if isinstance(value, str):
                if re.fullmatch(r"\d{2}/\d{2}", value):
                    month, year = map(int, value.split("/")); date = dt.datetime(2000 + year, month, 1)
                else:
                    try:
                        date = dt.datetime.combine(parse_date(value), dt.time())
                    except ReportError:
                        return value
            else:
                date = from_excel(float(value or 0))
            formats = {"dd/mm": "%d/%m", "mm/yy": "%m/%y", "dd/mm/yyyy": "%d/%m/%Y", "mmm": "%b", "mmmm": "%B", "dd/mm/yy": "%d/%m/%y"}
            if fmt in formats and hasattr(date, "strftime"):
                return date.strftime(formats[fmt])
            raise ReportError(f"Unsupported TEXT date format in the LIVE report: {fmt}")
        raise ReportError(f"Unsupported formula function in the LIVE report: {name}")

    def calculate(self, sheets: list[str]):
        for name in sheets:
            sheet = self.book.sheet(name)
            for (r, c), cell in sorted(sheet.cells.items()):
                if cell.formula:
                    try:
                        cell.value, cell.error = self.cell(sheet.name, r, c), False
                    except ExcelFault as fault:
                        # Preserve existing legacy errors; new reporting errors are
                        # checked before downloads are returned.
                        cell.value, cell.error = fault, True
            sheet.dirty = True


@lru_cache(maxsize=1)
def legacy_question_index():
    aliases = collections.defaultdict(list)
    for qid, provider, titles in QUESTION_ALIASES:
        for title in titles:
            aliases[title_key(title)].append((qid, provider))
    return aliases


def decode_questions(row: dict[str, str]) -> dict[str, str]:
    aliases = legacy_question_index()
    result = {}
    for header, value in row.items():
        match = re.match(r"^Q(\d+)\s*-\s*(.*)", header, re.I)
        if match:
            title = match[2]
            provider = "Amazon" if "amazon" in title.casefold() else "TNT" if "tnt" in title.casefold() else "generic"
            targets = [(int(match[1]), provider)]
        else:
            targets = aliases.get(title_key(header), [])
        for qid, provider in targets:
            key = f"{qid}:{provider}"
            if clean(value):
                if clean(result.get(key)) and result[key] != value:
                    raise ReportError(f"Conflicting answers for question {qid} in audit {row.get('internal_id', '')}.")
                result[key] = value
    return result


def question(record: dict, qid: int, provider: str = "generic") -> str:
    return record.get("_q", {}).get(f"{qid}:{provider}", "")


def coalesce_questions(record, ids, provider):
    values = list(dict.fromkeys(question(record, qid, provider) for qid in ids if clean(question(record, qid, provider))))
    if len(values) > 1:
        raise ReportError(f"Audit {record['internal_id']} has conflicting Amazon logo branch answers.")
    return values[0] if values else None


def canonical_fixture(value: str) -> str:
    value = unicodedata.normalize("NFKC", clean(value)).replace("\u200b", "")
    value = re.sub(r"\[[^\]]*\]", " ", value).replace("*", "")
    value = re.split(r"\s*-?\s*(?:positive(?:\s+site)?|pos)\s+confirmed\b", value, maxsplit=1, flags=re.I)[0]
    value = re.sub(r"\s+(?:TNT(?:\s+Sports?)?|Amazon(?:\s+Prime)?)\s*\d+\s*$", "", value, flags=re.I)
    value = re.sub(r"\b(?:vs\.?|versus)\b", "v", value, flags=re.I)
    value = re.sub(r"\s+v\.?\s+", " v ", value, flags=re.I)
    value = re.sub(r"\bvillareal\b", "villarreal", value, flags=re.I)
    return re.sub(r"\s+", " ", value).strip(" -").casefold()


def fixture_display(value: str) -> str:
    value = re.sub(r"\bv\b", "v", value.title(), flags=re.I)
    return re.sub(r"\b(fk|fc|psg)\b", lambda m: m[0].upper(), value, flags=re.I)


@dataclass(frozen=True)
class Game:
    date: str
    kickoff: str
    fixture: str

    @property
    def id(self):
        return hashlib.sha256(f"{self.date}|{self.kickoff}|{self.fixture}".encode()).hexdigest()[:16]

    @property
    def label(self):
        return f"{parse_date(self.date):%d/%m/%Y} {self.kickoff or 'time not specified'} — {fixture_display(self.fixture)}"


def game_from_visit_info(value: str) -> Game | None:
    match = re.match(r"^\s*(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2})\s+(.+)", value)
    if not match:
        return None
    fixture = canonical_fixture(match[3])
    if not re.search(r"\S\s+v\s+\S", fixture):
        return None
    hour, minute = map(int, match[2].split(":"))
    if hour > 23 or minute > 59:
        raise ReportError(f"Invalid kick-off time in Visit Info: {match[2]}")
    return Game(parse_date(match[1]).isoformat(), f"{hour:02}:{minute:02}", fixture)


@dataclass
class Plan:
    records: list[dict]
    games: dict[str, Game]
    skipped: int
    duplicates: int
    order_providers: dict[str, str]
    references: list[dict]
    cumulative_count: int
    amazon_count: int
    earliest_visit: str


def assign_games(records: list[dict], games: dict[str, Game]) -> dict[str, str | None]:
    assignments = {}
    for record in records:
        parsed = game_from_visit_info(record.get("visit_info", ""))
        if parsed and parsed.id in games:
            assignments[record["internal_id"]] = parsed.id
            continue
        day = record["_visit_date"]
        fixture = canonical_fixture(question(record, 1410))
        candidates = [key for key, game in games.items() if game.date == day and (not fixture or game.fixture == fixture)]
        if len(candidates) == 1:
            assignments[record["internal_id"]] = candidates[0]
        else:
            assignments[record["internal_id"]] = None
    return assignments


def reference_table(book: Book, name: str):
    if not any(n.strip() == name.strip() for n in book.paths):
        return {}
    sheet = book.sheet(name)
    result = {}
    for row in range(2, sheet.max_row + 1):
        code = clean(sheet.get(row, 1))
        if code:
            entry = dict(zip(("code", "priority", "name", "address", "city", "postcode"), (sheet.get(row, c) for c in range(1, 7))))
            entry["row"] = row
            previous = result.get(code)
            if previous and (clean(previous["priority"]) != clean(entry["priority"]) or clean(previous["city"]) != clean(entry["city"])):
                raise ReportError(f"The LIVE reference list contains conflicting entries for {code}.")
            result.setdefault(code, entry)
    return result


def prepare_plan(audit_bytes: bytes, previous_bytes: bytes) -> Plan:
    rows = read_csv_bytes(audit_bytes)
    if not rows:
        raise ReportError("The audit export has headers but no audits.")
    required = {"internal_id", "order_internal_id", "site_internal_id", "site_code", "primary_result", "visit_info"}
    missing = required - rows[0].keys()
    if missing:
        raise ReportError("Missing audit export columns: " + ", ".join(sorted(missing)))
    if not ({"date_of_visit", "date_of_visit_local"} & rows[0].keys()):
        raise ReportError("The export is missing its visit date column.")
    book = Book(previous_bytes)
    for name in (*INPUTS, *SUMMARIES, "Headline Figures", "Store Code"):
        if not any(n.strip() == name.strip() for n in book.paths):
            raise ReportError(f"The LIVE report is missing the '{name.strip()}' sheet.")
    cumulative = book.sheet("Cumulative Input")
    amazon = book.sheet("Amazon Data")
    cumulative_rows, amazon_rows = cumulative.audit_rows(), amazon.audit_rows()
    existing = {clean(sheet.get(row, 4)) for sheet in (cumulative, amazon) for row in sheet.audit_rows()}
    records, skipped, duplicates, seen = [], 0, 0, {}
    evidence = collections.defaultdict(set)
    for raw in rows:
        audit_id = clean(raw.get("internal_id"))
        if not audit_id:
            raise ReportError("An audit export row has no audit internal ID.")
        client = clean(raw.get("client_name"))
        if client and client.casefold() != "tnt sports":
            raise ReportError(f"The export contains another client ({client}). Export TNT Sports audits only.")
        q = decode_questions(raw)
        providers = {key.split(":", 1)[1] for key, value in q.items() if clean(value) and key.split(":", 1)[1] in ("TNT", "Amazon")}
        evidence[clean(raw.get("order_internal_id"))].update(providers)
        if audit_id in seen:
            if raw != seen[audit_id]:
                raise ReportError(f"Audit {audit_id} appears twice with different data in the export.")
            duplicates += 1
            continue
        seen[audit_id] = raw
        if audit_id in existing:
            skipped += 1
            continue
        if not clean(raw.get("site_code")):
            raise ReportError(f"Audit {audit_id} has no BTSCOM site code.")
        record = dict(raw)
        record["internal_id"] = audit_id
        record["_q"] = q
        record["_visit_date"] = parse_date(raw.get("date_of_visit_local") or raw.get("date_of_visit")).isoformat()
        record["_visit_time"] = clean(raw.get("time_of_visit_local") or raw.get("time_of_visit"))
        record["_direct_provider"] = next(iter(providers)) if len(providers) == 1 else ""
        record["_provider_conflict"] = len(providers) > 1
        records.append(record)
    order_providers = {order: next(iter(providers)) if len(providers) == 1 else "" for order, providers in evidence.items()}
    games = {}
    for record in records:
        game = game_from_visit_info(record.get("visit_info", ""))
        if game:
            games[game.id] = game
    # Auditors' fixture answers are a fallback when no scheduled fixture exists.
    for record in records:
        if game_from_visit_info(record.get("visit_info", "")):
            continue
        fixture = canonical_fixture(question(record, 1410))
        if re.search(r"\S\s+v\s+\S", fixture) and not any(g.date == record["_visit_date"] and g.fixture == fixture for g in games.values()):
            game = Game(record["_visit_date"], "", fixture); games[game.id] = game
    main_refs, amazon_refs = reference_table(book, "Store Code"), reference_table(book, "Amazon Store Code")
    references = {}
    for record in records:
        code = clean(record["site_code"])
        existing_ref = main_refs.get(code)
        fallback = amazon_refs.get(code, {})
        if existing_ref and clean(existing_ref.get("city")) and clean(existing_ref.get("priority")):
            continue
        reference = existing_ref or {}
        references.setdefault(code, {"code": code, "priority": clean(reference.get("priority")) or clean(fallback.get("priority")) or "Medium", "city": clean(reference.get("city")) or clean(fallback.get("city")), "name": record.get("site_name", ""), "address": record.get("site_address_1", ""), "postcode": record.get("site_post_code", ""), "site_id": clean(record["site_internal_id"]), "row": existing_ref["row"] if existing_ref else None})
    dates = [parse_date(cumulative.get(row, 15)).isoformat() for row in cumulative_rows]
    dates += [r["_visit_date"] for r in records]
    result = Plan(records, games, skipped, duplicates, order_providers, list(references.values()), len(cumulative_rows), len(amazon_rows), min(dates) if dates else dt.date.today().isoformat())
    book.archive.close(); book.loaded.clear()
    return result


def read_site_cities(content: bytes):
    rows = read_csv_bytes(content)
    if rows and not {"internal_id", "city"}.issubset(rows[0]):
        raise ReportError("The site export must contain internal_id and city columns.")
    by_id, codes = {}, collections.defaultdict(set)
    for row in rows:
        city = clean(row.get("city"))
        if city:
            site_id = clean(row.get("internal_id"))
            if site_id in by_id and by_id[site_id] != city:
                raise ReportError(f"The site export contains conflicting cities for {site_id}.")
            by_id[site_id] = city
            if clean(row.get("code")):
                codes[clean(row["code"])].add(city)
    return by_id, {code: next(iter(cities)) for code, cities in codes.items() if len(cities) == 1}


MAIN_QUESTIONS = [1, 303, 5475, 5476, 92, 228, 580, 5493, 5469, 579, 5470, 5471, 5472, 1410, 329, 3006, 3007, 5473, 1910, 1911, 3017, 3009, 3008, 314, 5642, 340, 320, 1919, 321, 585, 328, 333, 5491, 1411, 589, 3014, 1412, 3015, 3019, 3016, 337, 338, 345, 1912, 1913, 1914, 5474, 23, 614, 616]
PROVIDER_QUESTIONS = {321, 333, 1411, 589, 3014, 1412, 3015, 3016, 614}


def new_input_values(record: dict, provider: str, competition: str) -> dict[int, Any]:
    mapping = {1: "order_internal_id", 2: "order_schedule_type", 3: "client_name", 4: "internal_id", 5: "site_internal_id", 6: "responsibility", 7: "site_name", 8: "site_address_1", 9: "site_address_2", 10: "site_address_3", 11: "site_post_code"}
    result = {col: record.get(key) or None for col, key in mapping.items()}
    for col, field_name in ((12, "submitted_date"), (13, "approval_date")):
        result[col] = excel_date(record[field_name]) if clean(record.get(field_name)) else None
    result.update({14: record.get("item_to_order") or None, 15: excel_date(record["_visit_date"]), 16: record["_visit_time"] or None, 17: number_or_text(record.get("survey_score", "")), 18: record.get("primary_result") or None, 21: clean(record.get("site_code")), 24: competition})
    for col, qid in enumerate(MAIN_QUESTIONS, 26):
        result[col] = question(record, qid, provider if qid in PROVIDER_QUESTIONS else "generic") or None
    return result


def new_amazon_values(record: dict, competition: str) -> dict[int, Any]:
    standard = new_input_values(record, "Amazon", competition)
    result = {c: standard.get(c) for c in range(1, 12)}
    result.update({13: standard[12], 14: standard[13], 15: standard[14], 16: standard[15], 17: standard[16], 18: standard[17], 19: standard[18], 24: standard[21], 27: competition})
    for index, qid in enumerate(MAIN_QUESTIONS[:36], 30):
        # Amazon combines the two channel-logo branches into BL and BM.
        if qid == 1412:
            break
        result[index] = question(record, qid, "Amazon" if qid in PROVIDER_QUESTIONS else "generic") or None
    result[64] = coalesce_questions(record, (589, 1412), "Amazon")
    result[65] = coalesce_questions(record, (3014, 3015), "Amazon")
    for col, qid in enumerate([3019, 3016, 337, 338, 345, 1912, 1913, 1914, 5474, 23, 614, 616], 66):
        result[col] = question(record, qid, "Amazon" if qid in PROVIDER_QUESTIONS else "generic") or None
    return result


RANGE_PATTERN = re.compile(r"(?:(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_ ]*)!)?(?P<c1>\$?[A-Z]+)(?P<r1>\$?\d+):(?P<c2>\$?[A-Z]+)(?P<r2>\$?\d+)")


def extend_formula(formula: str, current_sheet: str, ends: dict[str, tuple[int, int]]) -> str:
    def replace(match):
        sheet = match["sheet"]
        sheet = sheet[1:-1].replace("''", "'") if sheet and sheet.startswith("'") else sheet or current_sheet
        normalized = next((name for name in ends if name.strip() == sheet.strip()), None)
        if normalized is None:
            return match[0]
        old, new = ends[normalized]
        if int(match["r2"].lstrip("$")) != old:
            return match[0]
        prefix = match[0][:match.start("r2") - match.start()]
        return prefix + ("$" if match["r2"].startswith("$") else "") + str(max(new, int(match["r1"].lstrip("$"))))
    return RANGE_PATTERN.sub(replace, formula)


def template_formulas(sheet: Sheet, row: int, start_col: int):
    return {col: cell.formula for (r, col), cell in sheet.cells.items() if r == row and col >= start_col and cell.formula}


def fill_input(sheet: Sheet, records: list[dict], provider: str, competitions: dict[str, str], start: int, ends: dict, replace_weekly=False):
    raw_end = 77 if provider == "Amazon" else 75
    templates = template_formulas(sheet, 3, raw_end + 1)
    for col in ((19, 20, 22, 23, 25) if provider != "Amazon" else (20, 21, 22, 23, 25, 26, 28, 29)):
        source = sheet.cells.get((3, col))
        if source and source.formula:
            templates[col] = source.formula
    if replace_weekly and 22 not in templates:
        exclusion = sheet.book.sheet("Cumulative Input").cells.get((3, 22))
        if exclusion and exclusion.formula:
            templates[22] = exclusion.formula.replace("'Cumulative Input'!", "'Weekly Input'!")
    if not any(col > raw_end for col in templates) and replace_weekly:
        # A prior week with no new TNT visits has no data-row formulas to copy.
        # Initialise the ordinary weekly helpers without changing any history.
        last = ends["Cumulative Input"][1]
        templates.update({76: '=VLOOKUP(U3,\'Store Code\'!A:B,2,FALSE)', 77: f'=1+COUNTIFS(\'Cumulative Input\'!$U$3:$U${last},U3,\'Cumulative Input\'!$O$3:$O${last},">"&O3)', 78: '=BX3&CB3', 80: '=IF(V3="","Include",V3)', 81: '=TEXT(O3,"dd/mm")', 82: '=CC3&R3', 83: '=BX3', 84: '=CC3&AA3', 85: '=TEXT(O3,"mm/yy")', 86: '=CG3&R3', 89: '=IF(Z3="Yes","Unsuccessful - Removed in advance","")', 90: '=IF(AE3="No","Unsuccessful - on the day","")', 91: '=IF(AE3="Yes","Successful","")', 92: '=IF(AN3="Yes"," TNT SPORT","")', 93: '=CK3&CL3&CM3&CN3', 94: '=COUNTIF(D:D,D3)'})
    if not any(col > raw_end for col in templates):
        raise ReportError(f"The {sheet.name} sheet has no formula template in row 3. Use the corrected LIVE report.")
    if replace_weekly:
        sheet.clear_from(3)
    for row, record in enumerate(records, start):
        competition = competitions[record["_game"]]
        values = new_amazon_values(record, competition) if provider == "Amazon" else new_input_values(record, provider, competition)
        for col in range(1, raw_end + 1):
            sheet.set(row, col, values.get(col), template_row=3)
        for col, formula in templates.items():
            expanded = extend_formula(formula, sheet.name, ends)
            translated = Translator(expanded, origin=f"{get_column_letter(col)}3").translate_formula(f"{get_column_letter(col)}{row}")
            sheet.set(row, col, formula=translated, template_row=3)
        if provider != "Amazon":
            auditor_col = 95 if replace_weekly else 101
            sheet.set(row, auditor_col, clean(record.get("auditor_internal_id")) or None, template_row=3)


def update_summary(sheet: Sheet, count: int, source_ends: dict):
    templates = {col: (cell.formula, cell.value) for (row, col), cell in sheet.cells.items() if row == 6}
    if not any(f for f, _ in templates.values()) and sheet.name == "Weekly Summary Table":
        source = sheet.book.sheet("Cumulative Summary ")
        templates = {}
        helpers = {"BY": "BX", "BZ": "BY", "CQ": "CO", "CW": "CQ"}
        for (r, col), cell in source.cells.items():
            if r != 6 or col == 11:
                continue
            formula = cell.formula
            if formula:
                formula = formula.replace("'Cumulative Input'!", "'Weekly Input'!")
                formula = re.sub(r"('Weekly Input'!)(\$?)([A-Z]+)(\$?\d+)", lambda m: m[1] + m[2] + helpers.get(m[3], m[3]) + m[4], formula)
            templates[col if col < 11 else col - 1] = (formula, cell.value)
    if not templates or not any(f for f, _ in templates.values()):
        raise ReportError(f"The {sheet.name.strip()} sheet is missing its formula template in row 6.")
    sheet.clear_from(count + 6)
    for row in range(6, count + 6):
        for col, (formula, value) in templates.items():
            existing = sheet.cells.get((row, col))
            if formula and existing is not None and existing.formula:
                continue
            if not formula and existing is not None and existing.value == value:
                continue
            new_formula = Translator(extend_formula(formula, sheet.name, source_ends), origin=f"{get_column_letter(col)}6").translate_formula(f"{get_column_letter(col)}{row}") if formula else None
            sheet.set(row, col, value=None if formula else value, formula=new_formula, template_row=6)
    sheet.filter_ref = f"A5:{get_column_letter(max(templates))}{max(count + 5, 5)}"


def update_headlines(book: Book, dates: list[dt.date], ends: dict):
    sheet = book.sheet("Headline Figures")
    # Row 5 includes a styled, intentionally blank percentage cell per date.
    # Copy both styles explicitly, including into existing unformatted blanks.
    total_styles = (sheet.cells[(5, 6)].style, sheet.cells[(5, 7)].style)
    count_formulas = {r: sheet.cells[(r, 6)].formula for r in range(5, 21)}
    percent_formulas = {r: sheet.cells[(r, 7)].formula for r in range(6, 21)}
    cumulative_percent = {r: sheet.cells[(r, 5)].formula for r in range(6, 21)}
    if not all(count_formulas.values()) or not all(percent_formulas.values()):
        raise ReportError("The Headline Figures sheet is missing its formula templates.")
    for (row, col), cell in list(sheet.cells.items()):
        if row >= 3 and col >= 6:
            sheet.set(row, col)
    letters = []
    def safe_percentage(formula):
        ratio = re.fullmatch(r"=(.+)/(\$?[A-Z]+\$?\d+)", formula)
        return f'=IF({ratio[2]}=0,"N/A",{formula[1:]})' if ratio else formula
    for index, day in enumerate(dates):
        col = 6 + 2 * index
        letter, pct = get_column_letter(col), get_column_letter(col + 1)
        letters.append(letter)
        sheet.set(3, col, day.strftime("%d/%m"), template_col=6)
        sheet.set(4, col, excel_date(day), template_col=6)
        sheet.set(4, col + 1, "%", template_col=7)
        sheet.extra_columns.extend([(col, 6), (col + 1, 7)])
        for offset, style in enumerate(total_styles):
            sheet.set(5, col + offset, template_col=6 + offset)
            sheet.cells[(5, col + offset)].style = style
        for row, formula in count_formulas.items():
            formula = Translator(extend_formula(formula, sheet.name, ends), origin=f"F{row}").translate_formula(f"{letter}{row}")
            sheet.set(row, col, formula=formula, template_col=6)
        for row, formula in percent_formulas.items():
            formula = Translator(formula, origin=f"G{row}").translate_formula(f"{pct}{row}")
            sheet.set(row, col + 1, formula=safe_percentage(formula), template_col=7)
    for row in range(5, 21):
        sheet.set(row, 4, formula="=SUM(" + ",".join(f"{letter}{row}" for letter in letters) + ")" if letters else "=0")
        if row in cumulative_percent:
            sheet.set(row, 5, formula=safe_percentage(cumulative_percent[row]))


def client_bytes(book: Book, counts: dict[str, int]) -> bytes:
    """A clean values-only package: raw input sheets/strings are not included."""
    selected = (["Invoicing"] if "Invoicing" in book.paths else []) + list(CLIENT_SHEETS)
    files = {}
    rels = ET.Element("Relationships", {"xmlns": PKG})
    workbook = ET.Element(Q("workbook"))
    views = ET.SubElement(workbook, Q("bookViews"))
    ET.SubElement(views, Q("workbookView"), {"activeTab": str(selected.index("Weekly Summary Table"))})
    sheets = ET.SubElement(workbook, Q("sheets"))
    content_types = ET.Element("Types", {"xmlns": CT})
    original_types = ET.fromstring(book.archive.read("[Content_Types].xml"))
    type_map = {e.get("PartName").lstrip("/"): e.get("ContentType") for e in original_types if e.tag.endswith("Override")}
    for entry in original_types:
        if entry.tag.endswith("Default"):
            ET.SubElement(content_types, "Default", dict(entry.attrib))
    def override(path, content_type):
        ET.SubElement(content_types, "Override", {"PartName": "/" + path, "ContentType": content_type})
    override("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    override("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
    files["xl/styles.xml"] = book.archive.read("xl/styles.xml")
    ET.SubElement(rels, "Relationship", {"Id": "rIdStyles", "Type": REL + "/styles", "Target": "styles.xml"})
    if "xl/theme/theme1.xml" in book.archive.namelist():
        files["xl/theme/theme1.xml"] = book.archive.read("xl/theme/theme1.xml")
        override("xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
        ET.SubElement(rels, "Relationship", {"Id": "rIdTheme", "Type": REL + "/theme", "Target": "theme/theme1.xml"})
    copied = set()
    def related_assets(path):
        if path in copied:
            return
        copied.add(path)
        files[path] = book.archive.read(path)
        if path in type_map:
            override(path, type_map[path])
        directory, basename = posixpath.split(path)
        relpath = directory + "/_rels/" + basename + ".rels"
        if relpath in book.archive.namelist():
            relations = ET.fromstring(book.archive.read(relpath))
            files[relpath] = book.archive.read(relpath)
            for relation in relations:
                if relation.get("TargetMode") != "External":
                    target = relation.get("Target")
                    related_assets(target.lstrip("/") if target.startswith("/") else posixpath.normpath(directory + "/" + target))
    for index, name in enumerate(selected, 1):
        sheet = book.sheet(name)
        path = f"xl/worksheets/sheet{index}.xml"
        state = "hidden" if name == "Invoicing" else "visible"
        ET.SubElement(sheets, Q("sheet"), {"name": sheet.name, "sheetId": str(index), "state": state, "{" + REL + "}id": f"rId{index}"})
        ET.SubElement(rels, "Relationship", {"Id": f"rId{index}", "Type": REL + "/worksheet", "Target": f"worksheets/sheet{index}.xml"})
        override(path, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        end_row = counts[name] + 5 if name in counts else None
        files[path] = sheet.xml(values_only=True, end_row=end_row)
        source_rel = posixpath.dirname(sheet.path) + "/_rels/" + posixpath.basename(sheet.path) + ".rels"
        if source_rel in book.archive.namelist():
            source_rels = ET.fromstring(book.archive.read(source_rel))
            kept = ET.Element("Relationships", {"xmlns": PKG})
            for relation in source_rels:
                if relation.get("Type", "").endswith("/drawing"):
                    ET.SubElement(kept, "Relationship", dict(relation.attrib))
                    target = relation.get("Target")
                    asset = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.dirname(sheet.path) + "/" + target)
                    related_assets(asset)
            if len(kept):
                files[f"xl/worksheets/_rels/sheet{index}.xml.rels"] = ET.tostring(kept, encoding="utf-8", xml_declaration=True)
    files["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
    files["xl/_rels/workbook.xml.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
    files["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)
    files["_rels/.rels"] = (f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>').encode()
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path, data in files.items():
            archive.writestr(path, data)
    return output.getvalue()


def default_folder_link(report_date: dt.date, earliest_visit: str) -> str:
    season_day = parse_date(earliest_visit)
    year = season_day.year if season_day.month >= 7 else season_day.year - 1
    return ("https://servelegalltd.sharepoint.com/:f:/r/sites/ServeLegal/Open%20Access/Media%20Compliance/TNT-SL-Data/Sent%20by%20SL/"
            f"Season%20{year % 100:02d}-{(year + 1) % 100:02d}/Reports/{report_date.year}/{report_date:%m.%y}")


def email_text(book: Book, new_tnt: list[dict], folder_link: str, report_date: dt.date):
    cumulative = book.sheet("Cumulative Input")
    rows = cumulative.audit_rows()
    periods = sorted({record["_visit_date"] for record in new_tnt})
    def rate(day=None):
        eligible = [row for row in rows if (day is None or parse_date(cumulative.get(row, 15)).isoformat() == day) and clean(cumulative.get(row, 81)).casefold() == "include" and clean(cumulative.get(row, 31)).casefold() == "yes"]
        matched = sum(clean(cumulative.get(row, 40)).casefold() == "yes" for row in eligible)
        return f"{math.floor(100 * matched / len(eligible) + 0.5)}%" if eligible else "N/A"
    lines = ["All,", "", f"The Serve Legal weekly report has been uploaded to the shared file location - {report_date:%m.%y} ({folder_link})", "", "As you'll see from the report, the infringement rates were:", ""]
    lines += [f"{parse_date(day):%d/%m} – {rate(day)}" for day in periods]
    if not periods:
        lines.append("No new TNT audits were added this week.")
    lines += ["", f"SEASON TOTAL – {rate()}", "", "Let me know if you have any issues."]
    return "\n".join(lines)


def merge_root(game_id: str, merges: dict[str, str]) -> str:
    seen = set()
    while merges.get(game_id):
        if game_id in seen:
            raise ReportError("Game grouping contains a loop. Keep one game as the final destination.")
        seen.add(game_id); game_id = merges[game_id]
    return game_id


def rich_email_html(text: str) -> str:
    paragraphs = []
    for block in text.split("\n\n"):
        lines = []
        for line in block.split("\n"):
            escaped_line = html.escape(line)
            if (re.fullmatch(r"\d{1,2}/\d{1,2}\s*[–-]\s*(?:\d+%|N/A)", line.strip())
                    or line.startswith("SEASON TOTAL") or line.strip() == "Kind Regards,"):
                escaped_line = "<strong>" + escaped_line + "</strong>"
            lines.append(escaped_line)
        escaped = "<br>".join(lines)
        escaped = re.sub(r'(\d{2}\.\d{2}) \((https?://[^\s<>]+)\)', lambda m: '<a href="' + m[2] + '" style="color:#0563c1;text-decoration:underline">' + m[1] + '</a>', escaped)
        paragraphs.append('<p style="margin:0 0 14px">' + escaped + '</p>')
    return '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000">' + "".join(paragraphs) + '</div>'


def email_copy_component(text: str) -> str:
    rich = rich_email_html(text)
    # JSON is placed in script data: escape '<' so a supplied URL cannot close it.
    rich_json = json.dumps(rich).replace("<", "\\u003c")
    return '''<!doctype html><html><head><meta charset="utf-8"><style>
    body{font-family:Arial,sans-serif;margin:0;padding:4px;color:#222}
    button{font-size:14px;padding:9px 15px;background:#fff;border:1px solid #aaa;
    border-radius:6px;cursor:pointer;margin-bottom:12px}#status{margin-left:12px;font-size:13px}
    #preview{padding:14px;border:1px solid #ddd;border-radius:6px;background:white}
    </style></head><body><button id="copy" type="button">Copy email for Outlook</button>
    <span id="status" role="status"></span><div id="preview" contenteditable="true" spellcheck="false">''' + rich + '''</div>
    <script>
    const htmlBody = ''' + rich_json + ''';
    const preview = document.getElementById('preview');
    function selectionCopy(){
      preview.focus();
      const range=document.createRange();range.selectNodeContents(preview);
      const selection=window.getSelection();selection.removeAllRanges();selection.addRange(range);
      const ok=document.execCommand('copy');
      if(ok)selection.removeAllRanges();return ok;
    }
    document.getElementById('copy').addEventListener('click',async()=>{
      const status=document.getElementById('status');
      // The synchronous selection route works inside Streamlit's iframe and
      // copies native HTML. Keep async ClipboardItem as the alternate route.
      let copied=false;try{copied=selectionCopy();}catch(error){}
      if(!copied){try{
        await navigator.clipboard.write([new ClipboardItem({
          'text/html':new Blob([preview.innerHTML],{type:'text/html'}),
          'text/plain':new Blob([preview.innerText],{type:'text/plain'})
        })]);copied=true;
      }catch(error){}}
      status.textContent=copied?'Copied — paste into Outlook.':
        'Select the email below and press Ctrl+C, then paste into Outlook.';
      if(!copied){const r=document.createRange();r.selectNodeContents(preview);
        const s=window.getSelection();s.removeAllRanges();s.addRange(r);}
    });
    </script></body></html>'''


def generate_reports(previous_bytes: bytes, plan: Plan, competitions: dict[str, str], *, report_date: dt.date, reference_overrides: dict[str, dict] | None = None, provider_overrides: dict[str, str] | None = None, game_overrides: dict[str, str] | None = None, custom_games: dict[str, Game] | None = None, merges: dict[str, str] | None = None, folder_link: str = ""):
    refs, providers, game_overrides, merges = reference_overrides or {}, provider_overrides or {}, game_overrides or {}, merges or {}
    games = dict(plan.games); games.update(custom_games or {})
    assignments = assign_games(plan.records, games)
    records = []
    for original in plan.records:
        record = dict(original)
        order = clean(record["order_internal_id"])
        provider = providers.get("override:" + order) or ("" if record["_provider_conflict"] else record["_direct_provider"]) or providers.get(order) or plan.order_providers.get(order)
        if provider not in ("TNT", "Amazon"):
            raise ReportError(f"Choose TNT or Amazon for order {order}.")
        record["_provider"] = provider
        game = game_overrides.get(record["internal_id"]) or assignments.get(record["internal_id"])
        if game not in games:
            raise ReportError(f"Assign audit {record['internal_id']} to a game.")
        game = merge_root(game, merges)
        if game not in games or not clean(competitions.get(game)):
            raise ReportError("Enter a competition for every game before generating.")
        record["_game"] = game
        records.append(record)
    for reference in plan.references:
        values = {**reference, **refs.get(reference["code"], {})}
        if not clean(values.get("city")):
            raise ReportError(f"A city is required for {reference['name']}. Upload the site export or fill the city in the reference table.")
        if not clean(values.get("priority")):
            raise ReportError(f"A priority is required for {reference['name']}.")
    book = Book(previous_bytes)
    cumulative, weekly, amazon = (book.sheet(name) for name in ("Cumulative Input", "Weekly Input", "Amazon Data"))
    old_cum, old_amazon, old_week = cumulative.audit_rows(), amazon.audit_rows(), weekly.audit_rows()
    tnt_records = [r for r in records if r["_provider"] == "TNT"]
    amazon_records = [r for r in records if r["_provider"] == "Amazon"]
    cum_end, amazon_end = max(old_cum, default=2), max(old_amazon, default=2)
    new_cum_end, new_amazon_end = cum_end + len(tnt_records), amazon_end + len(amazon_records)
    counts = {"Weekly Summary Table": len(tnt_records), "Cumulative Summary ": len(old_cum) + len(tnt_records), "Amazon Summary(Excluded)": len(old_amazon) + len(amazon_records)}
    ends = {"Cumulative Input": (cum_end, new_cum_end), "Amazon Data": (amazon_end, new_amazon_end), "Weekly Input": (max(old_week, default=2), len(tnt_records) + 2), "Weekly Summary Table": (len(old_week) + 5, len(tnt_records) + 5), "Cumulative Summary ": (len(old_cum) + 5, counts["Cumulative Summary "] + 5), "Amazon Summary(Excluded)": (len(old_amazon) + 5, counts["Amazon Summary(Excluded)"] + 5)}
    main_ref = book.sheet("Store Code")
    ref_end = max((r for (r, c), cell in main_ref.cells.items() if c == 1 and clean(cell.value)), default=1)
    for reference in plan.references:
        values = {**reference, **refs.get(reference["code"], {})}
        row = reference.get("row")
        if row:
            main_ref.set(row, 2, clean(values["priority"]))
            main_ref.set(row, 5, clean(values["city"]))
        else:
            ref_end += 1
            for col, key in enumerate(("code", "priority", "name", "address", "city", "postcode"), 1):
                main_ref.set(ref_end, col, clean(values.get(key)) or None, template_row=max(2, ref_end - 1))
    if plan.references:
        main_ref.filter_ref = f"A1:F{ref_end}"
    fill_input(cumulative, tnt_records, "TNT", competitions, cum_end + 1, ends)
    fill_input(amazon, amazon_records, "Amazon", competitions, amazon_end + 1, ends)
    fill_input(weekly, tnt_records, "TNT", competitions, 3, ends, replace_weekly=True)
    for sheet, last, end_col in ((weekly, len(tnt_records) + 2, "CQ"), (cumulative, new_cum_end, "DS"), (amazon, new_amazon_end, "CP")):
        sheet.filter_ref = f"A2:{end_col}{max(last, 2)}"
    for name in SUMMARIES:
        update_summary(book.sheet(name), counts[name], ends)
    dates = sorted({parse_date(cumulative.get(row, 15)) for row in cumulative.audit_rows()})
    update_headlines(book, dates, ends)
    calculation_sheets = [*INPUTS, *SUMMARIES, "Headline Figures"]
    for optional in ("=Checks", "Invoicing v2"):
        if optional in book.paths:
            book.sheet(optional); calculation_sheets.append(optional)
    for name in calculation_sheets:
        sheet = book.sheet(name)
        for (r, c), cell in list(sheet.cells.items()):
            if cell.formula:
                formula = extend_formula(cell.formula, sheet.name, ends)
                if formula != cell.formula:
                    sheet.set(r, c, cell.value, formula=formula)
    calculator = Calculator(book)
    # History is append-only. Reuse saved values for unchanged per-audit fields,
    # while recalculating ranks, duplicate checks, reference lookups and totals.
    # Newly written cells never qualify for this shortcut.
    static_input_columns = {
        "Cumulative Input": {22, 76, 81, 82, 83, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 97},
        "Amazon Data": {79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 92, 93, 94},
    }
    for name, columns in static_input_columns.items():
        for (r, col), cell in book.sheet(name).cells.items():
            if r >= 3 and col in columns and cell.formula and cell.trusted_cache:
                cell.calculated = True
    for name, changing_columns in (("Cumulative Summary ", {4, 6, 8}), ("Amazon Summary(Excluded)", {4})):
        for (r, col), cell in book.sheet(name).cells.items():
            if r >= 6 and col not in changing_columns and cell.formula and cell.trusted_cache:
                cell.calculated = True
    calculator.calculate(calculation_sheets)
    # Do not silently publish a newly introduced reporting error. Historical
    # cached errors from an uncorrected source remain the source's responsibility.
    for name in SUMMARIES:
        sheet = book.sheet(name)
        first_new = 6 if name == "Weekly Summary Table" else len(old_cum) + 6 if name == "Cumulative Summary " else len(old_amazon) + 6
        failures = [(f"{get_column_letter(c)}{r}", cell.value.code) for (r, c), cell in sheet.cells.items() if r >= first_new and isinstance(cell.value, ExcelFault)]
        if failures:
            raise ReportError(f"A new row in {name.strip()} has a calculation error at {failures[0][0]} ({failures[0][1]}). Check its store reference.")
    email = email_text(book, tnt_records, folder_link or default_folder_link(report_date, plan.earliest_visit), report_date)
    live = book.live_bytes()
    client = client_bytes(book, counts)
    validate_xlsx(live)
    validate_xlsx(client)
    book.archive.close()
    book.loaded.clear()
    return {"live": live, "client": client, "email": email, "counts": counts, "new_tnt": len(tnt_records), "new_amazon": len(amazon_records), "skipped": plan.skipped, "live_name": f"TNT Sports Weekly Report {report_date:%d.%m.%y} LIVE.xlsx", "client_name": f"TNT Sports Weekly Report {report_date:%d.%m.%y}.xlsx"}




def suggested_report_date(filename: str) -> dt.date:
    match = re.search(r"(\d{2})[._-](\d{2})[._-](\d{2,4})", filename)
    if match:
        day, month, year = map(int, match.groups())
        year += 2000 if year < 100 else 0
        try:
            return dt.date(year, month, day) + dt.timedelta(days=7)
        except ValueError:
            pass
    return dt.datetime.now(ZoneInfo("Europe/London")).date()


def main():
    import streamlit as st

    st.set_page_config(page_title="TNT Sports Report Generator", layout="wide")
    st.title("TNT Sports Weekly Report Generator")
    st.caption("Upload the audit export and previous LIVE report, assign competitions, then generate both reports.")
    left, right = st.columns(2)
    audit_file = left.file_uploader("Audit export", type=["csv"], help="Both old and new export formats are supported.")
    previous_file = right.file_uploader("Previous LIVE report", type=["xlsx"])
    if audit_file is None or previous_file is None:
        return
    audit_bytes, previous_bytes = audit_file.getvalue(), previous_file.getvalue()
    digest = hashlib.sha256(audit_bytes + previous_bytes).hexdigest()[:14]

    @st.cache_resource
    def generation_lock():
        return threading.Lock()

    @st.cache_data(show_spinner=False, max_entries=2)
    def cached_plan(audits, previous):
        with generation_lock():
            return prepare_plan(audits, previous)

    @st.cache_data(show_spinner=False, max_entries=2)
    def cached_cities(content):
        return read_site_cities(content)

    try:
        with st.spinner("Reading the uploaded files…"):
            plan = cached_plan(audit_bytes, previous_bytes)
    except (ReportError, ValueError, ET.ParseError, KeyError, zipfile.BadZipFile) as error:
        st.error(str(error)); return
    a, b, c = st.columns(3)
    a.metric("New audits", len(plan.records))
    b.metric("Already in the LIVE report", plan.skipped)
    c.metric("Games identified", len(plan.games))
    if plan.duplicates:
        st.info(f"Removed {plan.duplicates} repeated audit row(s) from the uploaded export.")
    if not plan.records:
        st.info("There are no new audit IDs in this export. Generating will produce an empty weekly TNT table and retain the cumulative history.")
    report_date = st.date_input("Report date", value=suggested_report_date(previous_file.name), key="report_date_" + digest)
    custom_key = "extra_games_" + digest
    custom_games = st.session_state.setdefault(custom_key, {})
    games = {**plan.games, **custom_games}
    ordered = sorted(games, key=lambda k: (games[k].date, games[k].kickoff, games[k].fixture))
    with st.expander("Add a game that could not be identified"):
        with st.form("extra_game_form_" + digest):
            first_date = min((parse_date(r["_visit_date"]) for r in plan.records), default=report_date)
            extra_date = st.date_input("Game date", value=first_date)
            extra_time = st.text_input("Kick-off", placeholder="20:00")
            extra_fixture = st.text_input("Game", placeholder="Home team v Away team")
            added = st.form_submit_button("Add game")
        if added:
            fixture = canonical_fixture(extra_fixture)
            if not re.search(r"\S\s+v\s+\S", fixture):
                st.error("Enter both teams with 'v' between them.")
            elif extra_time and (not re.fullmatch(r"(?:[01]?\d|2[0-3]):[0-5]\d", extra_time)):
                st.error("Enter the kick-off as HH:MM, or leave it blank.")
            else:
                kickoff = f"{int(extra_time.split(':')[0]):02}:{extra_time.split(':')[1]}" if extra_time else ""
                game = Game(extra_date.isoformat(), kickoff, fixture)
                st.session_state[custom_key][game.id] = game
                st.rerun()
    merges = {}
    possible_merges = [key for key in ordered if any(other != key and games[other].date == games[key].date for other in ordered)]
    if possible_merges:
        with st.expander("Combine game name variations"):
            st.caption("Use this only when two rows describe the same game. Venue notes and confirmation markers are already grouped automatically.")
            for key in possible_merges:
                others = [other for other in ordered if other != key and games[other].date == games[key].date]
                selected = st.selectbox(games[key].label, [""] + others, format_func=lambda value: games[value].label if value else "Keep as a separate game", key="merge_" + digest + key)
                if selected:
                    merges[key] = selected
    try:
        roots = {key: merge_root(key, merges) for key in ordered}
    except ReportError as error:
        st.error(str(error)); return
    reference_rows = []
    site_digest = "none"
    if plan.references:
        reference_rows = [dict(r) for r in plan.references]
        need_city = any(not clean(r["city"]) for r in reference_rows)
        site_digest = "none"
        if need_city:
            site_file = st.file_uploader("Site export", type=["csv"], help="Needed only for missing reference cities. The city column supplies the report's county field.", key="site_export_" + digest)
            if site_file is not None:
                try:
                    site_bytes = site_file.getvalue()
                    by_id, by_code = cached_cities(site_bytes)
                    site_digest = hashlib.sha256(site_bytes).hexdigest()[:8]
                    for reference in reference_rows:
                        if not clean(reference["city"]):
                            reference["city"] = by_id.get(reference["site_id"]) or by_code.get(reference["code"], "")
                except ReportError as error:
                    st.error(str(error)); return
    # Submit the entire configuration together, including the active editor cell.
    with st.form("report_configuration_" + digest):
        assignments = assign_games(plan.records, games)
        ambiguous = [r for r in plan.records if not assignments[r["internal_id"]]]
        game_overrides = {}
        if ambiguous:
            st.warning(f"Choose a game for {len(ambiguous)} audit(s) with insufficient or conflicting fixture details.")
            pending = pd.DataFrame([{"Audit ID": r["internal_id"], "Visit date": parse_date(r["_visit_date"]).strftime("%d/%m/%Y"), "Site": r.get("site_name", ""), "Visit Info": r.get("visit_info", ""), "Fixture answer": question(r, 1410), "Game": ""} for r in ambiguous])
            edited = st.data_editor(pending, hide_index=True, use_container_width=True, disabled=["Audit ID", "Visit date", "Site", "Visit Info", "Fixture answer"], column_config={"Game": st.column_config.SelectboxColumn("Game", options=[""] + [games[k].label for k in ordered], required=True)}, key="pending_games_" + digest + hashlib.sha256("".join(ordered).encode()).hexdigest()[:8])
            labels = {games[key].label: key for key in ordered}
            game_overrides = {row["Audit ID"]: labels.get(clean(row["Game"]), "") for row in edited.to_dict("records")}
        audit_counts = collections.Counter()
        for record in plan.records:
            game = game_overrides.get(record["internal_id"]) or assignments.get(record["internal_id"])
            if game:
                audit_counts[roots.get(game, game)] += 1
        if ambiguous and audit_counts:
            st.caption("Assigned audits: " + "; ".join(f"{games[key].label}: {audit_counts[key]}" for key in sorted(audit_counts)))
        root_games = sorted(set(roots.values()), key=lambda k: (games[k].date, games[k].kickoff, games[k].fixture))
        competitions = {}
        remembered_competitions = st.session_state.setdefault("competition_values_" + digest, {})
        for old_key, root in roots.items():
            if not remembered_competitions.get(root) and remembered_competitions.get(old_key):
                remembered_competitions[root] = remembered_competitions[old_key]
        if root_games:
            st.subheader("Competitions")
            st.caption("Enter the competition once for each game. It applies to all audits assigned to that game.")
            table = pd.DataFrame([{"Game ID": key, "Date": parse_date(games[key].date).strftime("%d/%m/%Y"), "Kick-off": games[key].kickoff, "Game": fixture_display(games[key].fixture), "Competition": remembered_competitions.get(key, "")} for key in root_games])
            editor_key = "competitions_" + digest + hashlib.sha256("".join(root_games).encode()).hexdigest()[:8]
            table = st.session_state.setdefault(editor_key + "_base", table)
            edited = st.data_editor(table, hide_index=True, use_container_width=True, disabled=["Game ID", "Date", "Kick-off", "Game"], column_config={"Game ID": None, "Competition": st.column_config.TextColumn("Competition", required=True, width="large")}, key=editor_key)
            competitions = {row["Game ID"]: clean(row["Competition"]) for row in edited.to_dict("records")}
            remembered_competitions.update(competitions)
        providers = {}
        orders = sorted({clean(r["order_internal_id"]) for r in plan.records})
        unresolved = []
        for order in orders:
            records = [r for r in plan.records if clean(r["order_internal_id"]) == order]
            if any(r["_provider_conflict"] or not (r["_direct_provider"] or plan.order_providers.get(order)) for r in records):
                unresolved.append(order)
        if unresolved:
            st.subheader("Broadcasters")
            st.caption("The exported answers do not identify a single broadcaster for every audit in these orders.")
            for order in unresolved:
                providers[order] = st.selectbox(f"Broadcaster for {order or '(order ID blank)'}", ["", "TNT", "Amazon"], key="provider_" + digest + order)
        if orders:
            with st.expander("Review detected broadcasters"):
                for order in orders:
                    if order in unresolved:
                        continue
                    detected = plan.order_providers.get(order) or "Mixed: identified per audit"
                    st.write(f"{order}: {detected}")
                    override = st.selectbox(f"Override {order}", ["Use detected broadcaster", "TNT", "Amazon"], key="provider_override_" + digest + order)
                    if override != "Use detected broadcaster":
                        providers["override:" + order] = override
        reference_overrides = {}
        if plan.references:
            st.subheader("Store references")
            st.caption("New priorities default to Medium. Override the priority or complete a missing city below.")
            priority_values = st.session_state.setdefault("priority_values_" + digest, {})
            table = pd.DataFrame([{"Code": r["code"], "Site": r["name"], "Priority": priority_values.get(r["code"], r["priority"]), "City": r["city"]} for r in reference_rows])
            editor_key = "references_" + digest + site_digest
            table = st.session_state.setdefault(editor_key + "_base", table)
            edited = st.data_editor(table, hide_index=True, use_container_width=True, disabled=["Code", "Site"], column_config={"Priority": st.column_config.TextColumn("Priority", required=True), "City": st.column_config.TextColumn("City", required=True)}, key=editor_key)
            reference_overrides = {r["Code"]: {"priority": clean(r["Priority"]), "city": clean(r["City"])} for r in edited.to_dict("records")}
            priority_values.update({code: values["priority"] for code, values in reference_overrides.items()})
        with st.expander("Email folder link"):
            folder = st.text_input("Report folder", value=default_folder_link(report_date, plan.earliest_visit), key="folder_v2_" + digest + report_date.isoformat())
        missing_competitions = [key for key in root_games if not competitions.get(key)]
        missing_assignments = [r for r in ambiguous if not game_overrides.get(r["internal_id"])]
        missing_providers = [order for order in unresolved if not providers.get(order)]
        missing_refs = [r for r in plan.references if not clean(reference_overrides.get(r["code"], {}).get("city", r["city"])) or not clean(reference_overrides.get(r["code"], {}).get("priority", r["priority"]))]
        ready = not (missing_competitions or missing_assignments or missing_providers or missing_refs)
        if not ready:
            todo = []
            if missing_competitions: todo.append("competitions")
            if missing_assignments: todo.append("game assignments")
            if missing_providers: todo.append("broadcasters")
            if missing_refs: todo.append("reference cities/priorities")
            st.info("Complete the " + ", ".join(todo) + " to generate the reports.")
        submitted = st.form_submit_button("Generate reports", type="primary")
    configuration = {"digest": digest, "date": report_date.isoformat(), "competitions": competitions, "providers": providers, "assignments": game_overrides, "merges": merges, "references": reference_overrides, "folder": folder}
    config_key = hashlib.sha256(json.dumps(configuration, sort_keys=True).encode()).hexdigest()
    output_key = "generated_reports_" + digest
    if submitted and ready:
        try:
            with st.spinner("Generating the LIVE report, client report and email…"):
                with generation_lock():
                    output = generate_reports(previous_bytes, plan, competitions, report_date=report_date, reference_overrides=reference_overrides, provider_overrides=providers, game_overrides=game_overrides, custom_games=custom_games, merges=merges, folder_link=folder)
            st.session_state[output_key] = (config_key, output)
        except (ReportError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as error:
            st.error(f"Could not generate the reports: {error}")
            st.session_state.pop(output_key, None)
    saved = st.session_state.get(output_key)
    if saved and saved[0] == config_key:
        output = saved[1]
        st.success(f"Reports generated: {output['new_tnt']} new TNT audits and {output['new_amazon']} new Amazon audits.")
        left, right = st.columns(2)
        left.download_button("Download updated LIVE report", output["live"], file_name=output["live_name"], mime=MIME_XLSX, on_click="ignore", use_container_width=True)
        right.download_button("Download client report", output["client"], file_name=output["client_name"], mime=MIME_XLSX, on_click="ignore", use_container_width=True)
        st.subheader("Email text")
        import streamlit.components.v1 as components
        st.caption("Copy the formatted email below, then paste it into a new Outlook message.")
        components.html(email_copy_component(output["email"]), height=600, scrolling=True)


# Old export titles corresponding to the IDs in new exports. Provider variants
# deliberately remain separate even when they reuse the same question ID.
QUESTION_ALIASES = [(1, 'generic', ['Did you identify during planning any reason why we should not visit this site?']),
 (303, 'generic', ['What is the reason for not visiting this site?']),
 (7583,
  'generic',
  ['Please describe why you were not comfortable auditing this site, or any incidents that occurred at the '
   'venue']),
 (5475, 'generic', ['Please provide a web link as supporting evidence for not visiting this site:']),
 (5476, 'generic', ['Please provide the BTSCOM Code for the duplicated venue:']),
 (92, 'generic', ['Please upload a site photo:']),
 (228, 'generic', ['Was the premises open for you to enter?']),
 (580, 'generic', ['If no, please detail why you could not enter?']),
 (5493, 'generic', ['If other, please detail the reason you could not enter:']),
 (5469, 'generic', ['Please detail the premises name for your Serve Legal audit:']),
 (579, 'generic', ['Please upload a screenshot of the address as displayed on Google:']),
 (5470, 'generic', ['Please enter the first line of the address for your Serve Legal audit:']),
 (5471, 'generic', ['Please enter the town/city name for your Serve Legal audit:']),
 (5472, 'generic', ['Please enter the postcode for your Serve Legal audit:']),
 (1410, 'generic', ['What live fixture were you sent to look for? (The Relevant Match)']),
 (329, 'generic', ['Was the Relevant Match showing on the TV screen(s) in the Premises?']),
 (3006, 'generic', ['What time did you enter the Premises?']),
 (3007, 'generic', ['What time did you leave the Premises?']),
 (5473, 'generic', ['Did you get a bank transaction/receipt, an internal photo, or both?']),
 (1910, 'generic', ['Please upload a screenshot of your bank transaction/receipt:']),
 (1911, 'generic', ['Please upload your internal photo:']),
 (3017, 'generic', ['Please describe a distinguishing feature, or the layout, of the Premises.']),
 (3009, 'generic', ['Please provide a description of a member of staff in the Premises:']),
 (3008, 'generic', ['Were there any customers present in the Premises?']),
 (314, 'generic', ['Did you see any sports advertising, either internally or externally?']),
 (5642, 'generic', ['What brand did you see advertised?']),
 (340, 'generic', ['If yes, please detail what you saw and where you saw it:']),
 (320, 'generic', ['How many TV screens did you see in the venue?']),
 (1919, 'generic', ['Were any of the TV screens on?']),
 (321,
  'TNT',
  ['Were any of the screens showing sport, sports news or other sports-related programs including those '
   'showing TNT Sport?']),
 (585, 'generic', ['Please select each sport that was being shown on the screens:']),
 (328, 'generic', ['If other, please state the sport being shown:']),
 (333, 'TNT', ['Was a TNT Sports channel logo visible on the TV screen(s) showing the Relevant Match?']),
 (5491, 'generic', ['What channel did you see displaying the Relevant Match?']),
 (1411, 'TNT', ["Did the TNT Sports channel logo contain the word 'Live'"]),
 (589, 'TNT', ['1. What TNT Sports channel logo was visible?']),
 (3014, 'TNT', ['1. If other, please detail the TNT Sports channel logo visible']),
 (1412, 'TNT', ['2. What TNT Sports channel logo was visible?']),
 (3015, 'TNT', ['2. If other, please detail the TNT Sports channel logo visible:']),
 (3019, 'generic', ['Where the channel logo was partly visible, please enter what you saw:']),
 (3016, 'TNT', ['In what position of the TV(s) was the TNT Sports channel logo seen?']),
 (337, 'generic', ['Was there a watermark visible on screen?']),
 (338, 'generic', ['If yes, what was the watermark?']),
 (345, 'generic', ['If other, please describe the watermark:']),
 (1912, 'generic', ['Were you able to take a photo of the screen showing the Relevant Match?']),
 (1913, 'generic', ['Please upload the photo of the screen showing the Relevant Match']),
 (1914, 'generic', ['Please upload a photo of your Grid']),
 (5474, 'generic', ['Do you wish to clarify any aspect of your report, or to explain anything unusual?']),
 (23,
  'generic',
  ['Please use this space to explain anything unusual about your visit or to clarify any detail of your '
   'report:']),
 (614,
  'TNT',
  ['Please confirm if this Premises was showing the Relevant Match or any other TNT Sports content']),
 (616,
  'generic',
  ['I confirm that I entered the Premises at the Address on the date and time stated in this report. I '
   'confirm that all the information that I have inputted on this report accurately reflects what I observed '
   'at the Premises and that I believe all facts in this report to be true']),
 (321,
  'Amazon',
  ['Were any of the screens showing sport, sports news or other sports-related programs including those '
   'showing Amazon Prime?']),
 (333,
  'Amazon',
  ['Was an Amazon Prime Sports channel logo visible on the TV screen(s) showing the Relevant Match?']),
 (1411, 'Amazon', ["Did the Amazon Prime channel logo contain the word 'Live'"]),
 (589, 'Amazon', ['What Amazon Prime channel logo was visible?']),
 (3014, 'Amazon', ['If other, please detail the Amazon Prime channel logo visible']),
 (1412, 'Amazon', ['What Amazon Prime channel logo was visible?']),
 (3015, 'Amazon', ['If other, please detail the Amazon Prime channel logo visible:']),
 (3016, 'Amazon', ['In what position of the TV(s) was the Amazon Prime channel logo seen?']),
 (614,
  'Amazon',
  ['Please confirm if this Premises was showing the Relevant Match or any other Amazon Prime content'])]



if __name__ == "__main__":
    main()
