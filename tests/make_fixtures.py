#!/usr/bin/env python3
"""Generate small workbooks used by tests/run_tests.ps1.

Requires: openpyxl, xlwt
"""
import os
import openpyxl
import xlwt
from openpyxl.workbook.defined_name import DefinedName


def make_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"

    ws["A1"] = 1
    ws["A2"] = 2
    ws["A3"] = 3
    ws["B1"] = "Hello"
    ws["B2"] = "World"

    ws["C1"] = "=A1+A2"
    ws["C2"] = "=SUM(A1:A3)"
    ws["C3"] = '=IF(A1>0,"yes","no")'

    ws["D1"] = '=A1&"x"'
    ws["D2"] = "=2^3"
    ws["D3"] = "=50%"

    ws["E1"] = "=ROUND(1.2345,2)"
    ws["E2"] = '=CONCATENATE(B1," ",B2)'
    ws["E3"] = "=LEFT(B1,2)"
    ws["E4"] = "=LEN(B1)"

    ws["F1"] = '=IFERROR(1/0,"err")'
    ws["F2"] = "=NA()"
    ws["F3"] = "=ISNUMBER(A1)"
    ws["F4"] = '=TEXT(1234.5,"#,##0.00")'

    ws["G1"] = "=DATE(2024,1,15)"
    ws["G2"] = "=YEAR(G1)"
    ws["G3"] = "=MONTH(G1)"
    ws["G4"] = "=DAY(G1)"

    ws["H1"] = '=VALUE("123")'
    ws["H2"] = '=TRIM("  a   b  ")'
    ws["H3"] = '=UPPER("abc")'
    ws["H4"] = '=MID("abcdef",2,3)'
    ws["H5"] = '=RIGHT("abc",2)'
    ws["H6"] = "=MOD(7,3)"
    ws["H7"] = "=ABS(-3)"
    ws["H8"] = "=POWER(2,3)"
    ws["H9"] = "=SQRT(9)"

    ws["I1"] = "=COUNT(A1:A3)"
    ws["I2"] = "=COUNTA(B1:B2)"
    ws["I3"] = "=AVERAGE(A1:A3)"
    ws["I4"] = "=MIN(A1:A3)"
    ws["I5"] = "=MAX(A1:A3)"
    ws["I6"] = "=AND(A1>0,A2>0)"
    ws["I7"] = "=OR(A1<0,A2>0)"
    ws["I8"] = "=NOT(A1>0)"
    ws["I9"] = "=ISBLANK(Z99)"
    ws["I10"] = "=ISERROR(1/0)"
    ws["I11"] = "=ISNA(NA())"

    ws["K1"] = 10
    ws["K2"] = 20
    ws["K3"] = 30

    ws["J1"] = "=SUM(A1:A3,10)"
    ws["J2"] = "=SUM(MyRange)"
    ws["J3"] = "=Other!A1+Data!A2"
    ws["J4"] = "='Other'!A1"
    ws["J5"] = '=IF(0,"a","b")'

    ws2 = wb.create_sheet("Other")
    ws2["A1"] = 10
    ws2["A2"] = 20

    wb.defined_names.add(DefinedName("MyRange", attr_text="Data!$A$1:$A$3"))
    wb.defined_names.add(DefinedName("OtherVal", attr_text="Other!$A$1"))
    wb.save(path)


def make_graph(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Graph"

    ws["A1"] = 1
    ws["A2"] = "=A1+1"
    ws["A3"] = "=A2*2"
    ws["B1"] = "=SUM(A1:A3)"
    ws["B2"] = "=B1+A3"
    ws["B3"] = "=B2-A1"

    # A deliberate circular chain.
    ws["C1"] = "=C2+1"
    ws["C2"] = "=C1+1"
    ws["C3"] = "=C1"

    ws["D1"] = "=IF(B2>10,B2,B3)"
    ws["D2"] = "=D1*2"
    ws["E1"] = "=SUM(A1:A3)+SUM(B1:B2)"
    ws["F1"] = "=E1+B3"

    ws2 = wb.create_sheet("Other")
    ws2["A1"] = "=Graph!B2"
    ws2["A2"] = "=Graph!C1"
    wb.save(path)


def make_types(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Types"

    # Text "02" (Excel's Text format) versus the number 2 in General,
    # and the number 2 displayed with a custom 00 format.
    ws["A1"] = "02"
    ws["A1"].number_format = "@"
    ws["A2"] = 2
    ws["A2"].number_format = "General"
    ws["A3"] = 2
    ws["A3"].number_format = "00"

    ws["A4"] = "=A1"
    ws["A5"] = "=A2"
    ws["A6"] = "=A3"
    ws["A7"] = '=TEXT(A2,"00")'

    ws["B1"] = '="02"'
    ws["B2"] = "=02"
    ws["B3"] = '=VALUE(A1)'
    ws["B4"] = "=ISTEXT(A1)"
    ws["B5"] = "=ISNUMBER(A2)"
    wb.save(path)


def make_fullcol(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "シート"

    ws["F1"] = "key"
    ws["G1"] = "name"
    ws["H1"] = "value"
    ws["F2"] = "a"
    ws["G2"] = "Alice"
    ws["H2"] = 10
    ws["F3"] = "b"
    ws["G3"] = "Bob"
    ws["H3"] = 20
    ws["F4"] = "c"
    ws["G4"] = "Carol"
    ws["H4"] = 30

    ws["L100"] = "b"
    wb.save(path)


def make_csv(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CSV"

    ws["A1"] = 'a,b"c'
    ws["A2"] = "line1\nline2"
    ws["A3"] = 1234.5
    ws["A3"].number_format = "#,##0.00"
    ws["A4"] = "02"
    ws["A4"].number_format = "@"
    ws["A5"] = 2
    ws["A6"] = 2
    ws["A6"].number_format = "00"
    wb.save(path)


def make_xls(path):
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Data")
    ws.write(0, 0, 1)
    ws.write(1, 0, 2)
    ws.write(2, 0, 3)
    ws.write(0, 1, "Hello")
    ws.write(1, 1, "World")
    ws2 = wb.add_sheet("Other")
    ws2.write(0, 0, 10)
    ws2.write(1, 0, 20)
    wb.save(path)


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "fixtures")
    os.makedirs(out, exist_ok=True)
    make_xlsx(os.path.join(out, "eval.xlsx"))
    make_graph(os.path.join(out, "graph.xlsx"))
    make_types(os.path.join(out, "types.xlsx"))
    make_fullcol(os.path.join(out, "fullcol.xlsx"))
    make_csv(os.path.join(out, "csv.xlsx"))
    make_xls(os.path.join(out, "eval.xls"))
    print("wrote tests/fixtures/eval.xlsx, graph.xlsx, types.xlsx, fullcol.xlsx, csv.xlsx and eval.xls")
