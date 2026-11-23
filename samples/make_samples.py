#!/usr/bin/env python3
"""Generate the 20,000 x 30 benchmark samples.

Requires: openpyxl, xlwt
    pip install openpyxl xlwt
"""
import argparse
import os
import time

import openpyxl
import xlwt


def cell_value(row: int, col: int) -> object:
    # row/col are 1-based.
    if col <= 10:
        return row * 1000 + col
    if col <= 20:
        return f"r{row}c{col}"
    if col <= 25:
        return 45000 + row
    return (row * col) % 1000


def make_xlsx(path: str, rows: int, cols: int) -> None:
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    for r in range(1, rows + 1):
        ws.append([cell_value(r, c) for c in range(1, cols + 1)])
    wb.save(path)


def make_xls(path: str, rows: int, cols: int) -> None:
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Data")
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.write(r - 1, c - 1, cell_value(r, c))
    wb.save(path)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", type=int, default=20000)
    ap.add_argument("--cols", type=int, default=30)
    ap.add_argument("--out", default=os.path.dirname(os.path.abspath(__file__)))
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    xlsx = os.path.join(args.out, f"sample_{args.rows}x{args.cols}.xlsx")
    xls = os.path.join(args.out, f"sample_{args.rows}x{args.cols}.xls")

    t = time.time()
    make_xlsx(xlsx, args.rows, args.cols)
    print(f"{xlsx}: {os.path.getsize(xlsx):,} bytes ({time.time() - t:.1f}s)")

    t = time.time()
    make_xls(xls, args.rows, args.cols)
    print(f"{xls}: {os.path.getsize(xls):,} bytes ({time.time() - t:.1f}s)")


if __name__ == "__main__":
    main()
