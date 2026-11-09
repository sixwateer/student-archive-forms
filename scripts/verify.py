#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验 generate.py 的输出：人数/学号完整性、表头页脚、A4 打印设置、列宽。"""
from __future__ import annotations

import argparse
import glob
import io
import json
import os
import zipfile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

INCH = 2.54


class Report:
    def __init__(self, path):
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        self.f = io.open(path, "w", encoding="utf-8")

    def __call__(self, *parts):
        line = " ".join(str(x) for x in parts)
        try:
            print(line)
        except UnicodeEncodeError:
            pass
        print(line, file=self.f)

    def close(self):
        self.f.close()


def main():
    ap = argparse.ArgumentParser(description="校验档案表输出")
    ap.add_argument("--source", required=True, help="学生名册（与生成时一致）")
    ap.add_argument("--template", required=True, help="模板 xlsx")
    ap.add_argument("--out", required=True, help="生成时的输出目录")
    ap.add_argument("--header-row", type=int, default=3)
    ap.add_argument("--id-column", default="学号")
    ap.add_argument("--margin-cm", type=float, default=2.0)
    ap.add_argument("--footer-gap", type=int, default=3, help="与生成时保持一致")
    ap.add_argument("--include-zip", action="store_true", help="同时校验目录下的 zip 完整性")
    ap.add_argument("--ignore-filter", action="store_true",
                    help="忽略输出目录里的 _筛选条件.json，仍按 --source 的全量名册校验")
    ap.add_argument("--log", default="")
    a = ap.parse_args()
    rep = Report(a.log or os.path.join(a.out, "_校验报告.txt"))
    try:
        run(a, rep)
    finally:
        rep.close()


def run(a, rep):
    # generate.py 若用了筛选，会写出 _筛选条件.json + 筛选后名册；
    # 此时自动改按筛选后的学生核对，避免把「没生成的人」误报成缺失。
    meta_path = os.path.join(a.out, "_筛选条件.json")
    if not a.ignore_filter and os.path.isfile(meta_path):
        try:
            with io.open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
            sub = meta.get("subset") or ""
            if meta.get("enabled") and sub and os.path.isfile(sub):
                rep("检测到筛选生成，改用筛选后名册校验:", sub)
                rep("筛选条件:", "；".join(meta.get("conditions") or []))
                a.source, a.header_row = sub, 1
        except (ValueError, OSError, TypeError):
            pass

    df = pd.read_excel(a.source, sheet_name=0, header=a.header_row - 1, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df[a.id_column].notna()].copy()
    src_ids = set(df[a.id_column].astype(str).str.strip())
    src_classes = df["班级"].astype(str).str.strip().nunique() if "班级" in df.columns else None

    tpl = load_workbook(a.template).worksheets[0]
    header = [tpl.cell(row=1, column=c).value for c in range(1, tpl.max_column + 1)]
    header = [h for h in header if h is not None]
    # 模板页脚：跳过可能存在的示例数据行（按最大行间隔切分）
    raw = [(c.column, c.row, c.value) for row in tpl.iter_rows(min_row=2)
           for c in row if c.value is not None]
    rows = sorted({r for _, r, _ in raw})
    start = rows[0] if rows else None
    if len(rows) > 1:
        gap, idx = max((rows[i + 1] - rows[i], i) for i in range(len(rows) - 1))
        if gap >= a.footer_gap:
            start = rows[idx + 1]
    footer_cells = [(c, r, v) for c, r, v in raw if r >= start]
    footer_first = footer_cells[0][2] if footer_cells else None
    footer_col = footer_cells[0][0] if footer_cells else 4

    files = sorted(glob.glob(os.path.join(a.out, "*.xlsx")))
    rep("=" * 70)
    rep("校验报告 | 源表 %d 人 | 输出目录 %s" % (len(src_ids), a.out))
    rep("=" * 70)

    problems = []
    seen_ids = []
    n_sheet = n_row = 0
    for path in files:
        wb = load_workbook(path)
        base = os.path.basename(path)
        f_sheets = f_rows = 0
        for sn in wb.sheetnames:
            ws = wb[sn]
            n_sheet += 1
            f_sheets += 1
            hdr = [ws.cell(row=1, column=c).value for c in range(1, len(header) + 1)]
            if hdr != header:
                problems.append(f"{base}/{sn}: 表头不一致 {hdr}")

            n = 0
            r = 2
            while isinstance(ws.cell(row=r, column=1).value, int):
                n += 1
                r += 1
            n_row += n
            f_rows += n
            for i in range(n):
                rr = i + 2
                if ws.cell(row=rr, column=1).value != i + 1:
                    problems.append(f"{base}/{sn}: 序号错乱 @row{rr}")
                    break
                seen_ids.append(str(ws.cell(row=rr, column=3).value))
                if ws.cell(row=rr, column=6).value is not None or ws.cell(row=rr, column=7).value is not None:
                    problems.append(f"{base}/{sn}: 是否存档/备注非空 @row{rr}")
                    break
            if footer_first is not None and ws.cell(row=n + 2, column=footer_col).value != footer_first:
                problems.append(f"{base}/{sn}: 页脚位置异常")

            ps, pm = ws.page_setup, ws.page_margins
            if ps.paperSize != 9 or ps.orientation != "portrait" or ps.fitToWidth != 1:
                problems.append(f"{base}/{sn}: A4/缩放设置异常")
            if ws.sheet_properties.pageSetUpPr is None or not ws.sheet_properties.pageSetUpPr.fitToPage:
                problems.append(f"{base}/{sn}: 未启用 fitToPage")
            for attr in ["left", "right", "top", "bottom"]:
                if abs(getattr(pm, attr) * INCH - a.margin_cm) > 0.02:
                    problems.append(f"{base}/{sn}: 边距 {attr}={getattr(pm, attr)*INCH:.2f}cm")
            if (ws.print_title_rows or "").replace("$", "") != "1:1":
                problems.append(f"{base}/{sn}: 未设重复表头")
            if not ws.print_area:
                problems.append(f"{base}/{sn}: 未设打印区域")

        w = sum(wb.worksheets[0].column_dimensions[get_column_letter(c)].width or 8.43
                for c in range(1, len(header) + 1))
        rep(f"  {base:<52} sheet={f_sheets:>3} 人={f_rows:>5} 列宽≈{w*0.185*1.05:.1f}cm")

    rep("-" * 70)
    rep(f"文件={len(files)} sheet={n_sheet} 学生行={n_row} | 源表 {len(src_ids)} 人")
    if src_classes:
        rep("班级数一致:", n_sheet == src_classes, f"(输出 {n_sheet} / 源表 {src_classes})")
    miss = src_ids - set(seen_ids)
    extra = set(seen_ids) - src_ids
    dup = len(seen_ids) - len(set(seen_ids))
    rep("人数一致:", n_row == len(src_ids), "| 学号缺失:", len(miss), "多余:", len(extra), "重复:", dup)
    if miss:
        rep("  缺失样例:", list(miss)[:10])
    if extra:
        rep("  多余样例:", list(extra)[:10])

    if a.include_zip:
        for z in sorted(glob.glob(os.path.join(a.out, "*.zip"))):
            with zipfile.ZipFile(z) as f:
                rep(f"  ZIP {os.path.basename(z)}: {len(f.namelist())} 个文件, 完整性 {'OK' if f.testzip() is None else 'FAIL'}")

    if miss or extra or dup or n_row != len(src_ids) or problems:
        rep("\n异常项 (%d):" % (len(problems) + len(miss) + len(extra)))
        for x in problems[:40]:
            rep("  -", x)
    else:
        rep("\n全部校验通过：人数、学号、表头、页脚、A4 打印设置均一致，异常项 0。")


if __name__ == "__main__":
    main()
