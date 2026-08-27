"""CSV · XLSX 읽기 테스트 (네트워크 불필요)."""

import datetime
import os
import tempfile
import unittest
import zipfile

from pqr.tabular import TableError, excel_serial_to_date, read_csv, read_table, read_xlsx

SHEET_XML = """<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>
<row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2"><v>99.4</v></c><c r="C2"><v>45700</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>HP-102</t></is></c><c r="B3"><v>101.2</v></c><c r="C3"><v>45701</v></c></row>
</sheetData></worksheet>"""

SHARED = """<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
<si><t>제품코드</t></si><si><t>결과값</t></si><si><t>제조일</t></si><si><t>HP-101</t></si></sst>"""

WORKBOOK = """<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="데이터" sheetId="1" r:id="rId1"/></sheets></workbook>"""

RELS = """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""


def make_xlsx(path):
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", WORKBOOK)
        archive.writestr("xl/_rels/workbook.xml.rels", RELS)
        archive.writestr("xl/sharedStrings.xml", SHARED)
        archive.writestr("xl/worksheets/sheet1.xml", SHEET_XML)


class TabularTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def _write(self, name, text, encoding="utf-8"):
        path = os.path.join(self.dir, name)
        with open(path, "wb") as handle:
            handle.write(text.encode(encoding))
        return path

    def test_csv_utf8_bom(self):
        path = self._write("a.csv", "제품코드,결과값\nHP-101,99.4\n", "utf-8-sig")
        rows = read_csv(path)
        self.assertEqual(rows, [{"제품코드": "HP-101", "결과값": "99.4"}])

    def test_csv_cp949(self):
        path = self._write("b.csv", "제품코드,결과값\nHP-101,99.4\n", "cp949")
        self.assertEqual(read_csv(path)[0]["제품코드"], "HP-101")

    def test_csv_semicolon_and_blank_lines(self):
        path = self._write("c.csv", "코드;값\n\nHP-1;3\n\n")
        self.assertEqual(read_csv(path), [{"코드": "HP-1", "값": "3"}])

    def test_csv_short_row_is_padded(self):
        path = self._write("d.csv", "a,b,c\n1,2\n")
        self.assertEqual(read_csv(path), [{"a": "1", "b": "2", "c": ""}])

    def test_xlsx_shared_and_inline_strings(self):
        path = os.path.join(self.dir, "x.xlsx")
        make_xlsx(path)
        rows = read_xlsx(path)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["제품코드"], "HP-101")
        self.assertEqual(rows[1]["제품코드"], "HP-102")
        self.assertEqual(rows[0]["결과값"], "99.4")

    def test_read_table_dispatches_by_suffix(self):
        path = os.path.join(self.dir, "y.xlsx")
        make_xlsx(path)
        self.assertEqual(len(read_table(path)), 2)
        with self.assertRaises(TableError):
            read_table(os.path.join(self.dir, "z.pdf"))

    def test_corrupt_xlsx_reports_clearly(self):
        path = self._write("broken.xlsx", "이건 엑셀이 아닙니다")
        with self.assertRaises(TableError):
            read_table(path)

    def test_excel_serial_to_date(self):
        self.assertEqual(excel_serial_to_date("45700"), datetime.date(2025, 2, 12))
        self.assertIsNone(excel_serial_to_date("abc"))
        self.assertIsNone(excel_serial_to_date("0.5"))


if __name__ == "__main__":
    unittest.main()
