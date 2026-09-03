# -*- coding: utf-8 -*-
"""Cpk 계산 파일(.xls) 을 서식 그대로 채운다.

이 서식은 수식과 꺾은선 그래프가 들어 있는 .xls 다. .xlsx 로 바꿔서 openpyxl 로 저장하면
그래프가 사라지므로, 엑셀(Windows) 이나 LibreOffice 를 시켜 원래 형식 그대로 채운다.
서식의 그래프는 11개 Lot 까지만 잡혀 있어 Lot 수에 맞춰 범위를 늘린다.
"""
import os
import re
import shutil
import subprocess
import sys
import time

FIRST_DATA_ROW = 10          # B10 부터 결과값
ROWS = 35                    # B10 ~ B44
HEAD_ROW = 9                 # 그래프가 계열 이름으로 쓰는 머리 행
CALC_COL_FIRST = "R"         # 그래프가 쓰는 계산 열
CALC_COL_LAST = {True: "V", False: "T"}      # 양쪽 규격 / 한쪽 규격


class FillError(Exception):
    pass


def _last_row(n):
    return HEAD_ROW + max(n, 1)


# ---------------------------------------------------------------- Excel (COM)
def _with_excel(src, dst, cells, values):
    try:
        import win32com.client
    except ImportError:
        return False
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(src))
        ws = wb.Worksheets(1)
        for ref, value in cells.items():
            ws.Range(ref).Value = value
        for i in range(ROWS):
            ws.Cells(FIRST_DATA_ROW + i, 2).Value = values[i] if i < len(values) else ""
        last = _last_row(len(values))
        for i in range(1, ws.ChartObjects().Count + 1):
            chart = ws.ChartObjects(i).Chart
            for j in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(j)
                series.Formula = re.sub(
                    r"\$([A-Z]+)\$(\d+):\$([A-Z]+)\$\d+",
                    lambda m: "$%s$%s:$%s$%d" % (m.group(1), m.group(2), m.group(3), last),
                    series.Formula)
        wb.Application.CalculateFull()
        wb.SaveAs(os.path.abspath(dst), FileFormat=56)        # xlExcel8 (.xls)
        wb.Close(False)
    finally:
        excel.Quit()
    return os.path.isfile(dst)


# ----------------------------------------------------------- LibreOffice(UNO)
def _uno_desktop(port):
    import uno
    from com.sun.star.connection import NoConnectException
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    url = ("uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port)
    for _ in range(30):
        try:
            ctx = resolver.resolve(url)
        except NoConnectException:
            time.sleep(1)
            continue
        return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return None


def _soffice_listener(port):
    from . import convert
    exe = convert._soffice()
    if not exe:
        return None
    return subprocess.Popen(
        [exe, "--headless", "--norestore", "--invisible",
         "--accept=socket,host=127.0.0.1,port=%d;urp;" % port],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _with_uno(src, dst, cells, values, port=2103):
    try:
        import uno
        from com.sun.star.beans import PropertyValue
        from com.sun.star.table import CellRangeAddress
    except ImportError:
        return False
    desktop = _uno_desktop(port)
    started = None
    if desktop is None:
        started = _soffice_listener(port)
        if started is None:
            return False
        desktop = _uno_desktop(port)
    if desktop is None:
        return False

    def prop(name, value):
        p = PropertyValue()
        p.Name, p.Value = name, value
        return p

    # 한글·공백이 든 경로는 LibreOffice 의 형식 인식이 실패한다 — ASCII 이름으로 다룬다.
    work = os.path.join(os.path.dirname(os.path.abspath(dst)), "_cpk_work")
    os.makedirs(work, exist_ok=True)
    plain_src = os.path.join(work, "in.xls")
    plain_dst = os.path.join(work, "out.xls")
    shutil.copyfile(src, plain_src)
    try:
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(plain_src), "_blank", 0,
                                           (prop("Hidden", True),))
        sheet = doc.Sheets.getByIndex(0)
        for ref, value in cells.items():
            cell = sheet.getCellRangeByName(ref)
            if isinstance(value, (int, float)):
                cell.setValue(float(value))
            else:
                cell.setString("" if value is None else str(value))
        for i in range(ROWS):
            cell = sheet.getCellByPosition(1, FIRST_DATA_ROW - 1 + i)
            if i < len(values):
                cell.setValue(float(values[i]))
            else:
                cell.setString("")
        wide = "Nominal" in sheet.getCellRangeByName("O5").getString()
        last = _last_row(len(values)) - 1                     # 0 부터 세는 행 번호
        for name in sheet.Charts.ElementNames:
            chart = sheet.Charts.getByName(name)
            areas = []
            for c1, c2 in ((1, 1), (17, 21 if wide else 19)):
                a = CellRangeAddress()
                a.Sheet, a.StartColumn, a.StartRow = 0, c1, HEAD_ROW - 1
                a.EndColumn, a.EndRow = c2, last
                areas.append(a)
            chart.Ranges = tuple(areas)
        doc.calculateAll()
        if os.path.exists(plain_dst):
            os.remove(plain_dst)
        doc.storeToURL(uno.systemPathToFileUrl(plain_dst),
                       (prop("FilterName", "MS Excel 97"),))
        doc.close(False)
        shutil.copyfile(plain_dst, dst)
    finally:
        shutil.rmtree(work, ignore_errors=True)
        if started is not None:
            started.terminate()
    return os.path.isfile(dst)


def fill(src, dst, cells, values):
    """src(.xls) 서식에 cells 와 결과값을 채워 dst(.xls) 로 저장한다. 쓴 방법을 돌려준다."""
    if sys.platform == "win32" and _with_excel(src, dst, cells, values):
        return "excel"
    if _with_uno(src, dst, cells, values):
        return "soffice"
    raise FillError("Cpk 계산 파일을 채우려면 Excel 또는 LibreOffice 가 필요합니다.")
