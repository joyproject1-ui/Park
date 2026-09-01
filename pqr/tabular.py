"""CSV · XLSX 읽기 (표준 라이브러리만 사용).

사내망에서 별도 패키지 설치 없이 동작해야 하므로 openpyxl·pandas 대신
zipfile + xml.etree 로 xlsx 를 직접 읽습니다.
"""

import csv
import datetime as _dt
import io
import re
import zipfile
import xml.etree.ElementTree as ET

_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# 사내 파일은 UTF-8(BOM) 또는 CP949 로 저장되는 경우가 대부분입니다.
_ENCODINGS = ("utf-8-sig", "utf-8", "cp949", "euc-kr", "latin-1")

_EXCEL_EPOCH = _dt.date(1899, 12, 30)


class TableError(Exception):
    """읽을 수 없는 파일이거나 표 구조가 아닐 때."""


def excel_serial_to_date(value):
    """엑셀 날짜 일련번호를 date 로 바꿉니다. 변환할 수 없으면 None."""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    # 1900-01-01 ~ 2100 년 범위를 벗어나는 값은 날짜가 아니라고 봅니다.
    if not (1 <= number <= 80000):
        return None
    return _EXCEL_EPOCH + _dt.timedelta(days=int(number))


def _decode(raw):
    for encoding in _ENCODINGS:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise TableError("파일 인코딩을 판별하지 못했습니다 (UTF-8 또는 CP949 로 저장하세요)")


def _sniff_delimiter(sample):
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        return ","


def read_csv(path):
    with open(path, "rb") as handle:
        text = _decode(handle.read())
    if not text.strip():
        return []
    delimiter = _sniff_delimiter(text[:4096])
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = [row for row in reader if any(cell.strip() for cell in row)]
    if not rows:
        return []
    header = [cell.strip() for cell in rows[0]]
    return [_row_to_dict(header, row) for row in rows[1:]]


def _row_to_dict(header, row):
    record = {}
    for index, name in enumerate(header):
        if not name:
            continue
        record[name] = row[index].strip() if index < len(row) else ""
    return record


def _column_index(reference):
    """'BC12' -> 54 (0-based 열 번호)."""
    letters = re.match(r"([A-Z]+)", reference or "")
    if not letters:
        return None
    index = 0
    for char in letters.group(1):
        index = index * 26 + (ord(char) - 64)
    return index - 1


def _shared_strings(archive):
    try:
        raw = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    strings = []
    for item in root.findall("m:si", _NS):
        strings.append("".join(node.text or "" for node in item.iter("{%s}t" % _NS["m"])))
    return strings


def _sheet_targets(archive):
    """워크북에 정의된 순서대로 시트 XML 경로를 돌려줍니다."""
    book = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    by_id = {}
    for rel in rels.findall("pr:Relationship", _NS):
        by_id[rel.get("Id")] = rel.get("Target")
    targets = []
    for sheet in book.findall("m:sheets/m:sheet", _NS):
        rid = sheet.get("{%s}id" % _NS["r"])
        target = by_id.get(rid)
        if not target:
            continue
        if not target.startswith("/"):
            target = "xl/" + target.lstrip("./")
        targets.append((sheet.get("name") or "", target.lstrip("/")))
    return targets


def read_xlsx(path, sheet=None):
    """첫 번째 시트(또는 이름이 일치하는 시트)를 dict 목록으로 읽습니다."""
    try:
        archive = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise TableError("xlsx 파일이 아니거나 손상되었습니다: %s" % path)
    with archive:
        strings = _shared_strings(archive)
        targets = _sheet_targets(archive)
        if not targets:
            raise TableError("시트를 찾을 수 없습니다: %s" % path)
        target = targets[0][1]
        if sheet:
            for name, candidate in targets:
                if name == sheet:
                    target = candidate
                    break
            else:
                raise TableError("시트 '%s' 를 찾을 수 없습니다" % sheet)
        root = ET.fromstring(archive.read(target))

    grid = []
    for row_node in root.findall("m:sheetData/m:row", _NS):
        cells = {}
        for cell in row_node.findall("m:c", _NS):
            index = _column_index(cell.get("r"))
            if index is None:
                continue
            cells[index] = _cell_value(cell, strings)
        if not cells:
            continue
        width = max(cells) + 1
        grid.append([cells.get(i, "") for i in range(width)])

    grid = [row for row in grid if any(str(cell).strip() for cell in row)]
    if not grid:
        return []
    header = [str(cell).strip() for cell in grid[0]]
    return [_row_to_dict(header, [str(cell) for cell in row]) for row in grid[1:]]


def _cell_value(cell, strings):
    kind = cell.get("t")
    if kind == "inlineStr":
        node = cell.find("m:is", _NS)
        if node is None:
            return ""
        return "".join(part.text or "" for part in node.iter("{%s}t" % _NS["m"]))
    value_node = cell.find("m:v", _NS)
    if value_node is None or value_node.text is None:
        return ""
    text = value_node.text
    if kind == "s":
        try:
            return strings[int(text)]
        except (ValueError, IndexError):
            return ""
    if kind == "b":
        return "TRUE" if text == "1" else "FALSE"
    return text


def read_xls(path, sheet=None):
    """구형 엑셀(.xls) — ERP 출력물이 이 형식입니다.

    표준 라이브러리로는 읽을 수 없어 xlrd 가 있을 때만 동작합니다.
    ERP 가 만든 .xls 에는 CODEPAGE 레코드가 없어서 xlrd 가 iso-8859-1 로 읽습니다.
    그러면 '나조린점안액' 이 '³ªÁ¶¸°Á¡¾È¾×' 으로 나옵니다 — cp949 로 못박아 읽습니다.
    """
    try:
        import xlrd
    except ImportError:
        raise TableError(
            "구형 엑셀(.xls)은 xlrd 가 있어야 읽습니다. 엑셀에서 '다른 이름으로 저장 → "
            "Excel 통합 문서(*.xlsx)' 로 바꿔 올리면 추가 설치 없이 읽습니다: %s" % path)
    try:
        book = xlrd.open_workbook(path, encoding_override="cp949")
    except Exception as error:                       # xlrd 는 형식별로 예외가 제각각입니다
        raise TableError("xls 파일을 읽지 못했습니다: %s (%s)" % (path, error))
    names = book.sheet_names()
    target = book.sheet_by_name(sheet) if sheet in names else book.sheet_by_index(0)
    rows = []
    for index in range(target.nrows):
        rows.append(["" if value is None else str(value).strip()
                     for value in target.row_values(index)])
    rows = [row for row in rows if any(cell for cell in row)]
    if not rows:
        return []
    header = [str(cell).strip() for cell in rows[0]]
    return [_row_to_dict(header, row) for row in rows[1:]]


def read_table(path, sheet=None):
    """확장자를 보고 CSV · XLSX · XLS 로 읽습니다."""
    lowered = str(path).lower()
    if lowered.endswith((".xlsx", ".xlsm")):
        return read_xlsx(path, sheet=sheet)
    if lowered.endswith(".xls"):
        return read_xls(path, sheet=sheet)
    if lowered.endswith((".csv", ".tsv", ".txt")):
        return read_csv(path)
    raise TableError("지원하지 않는 형식입니다 (csv · xlsx · xls 만 지원): %s" % path)
