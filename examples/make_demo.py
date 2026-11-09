#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成完全虚构的演示数据：一份学生名册 + 一份纸质表模板。

姓名、学号、学院、班级全部是编造的，仅用于试跑本工具，不含任何真实学生信息。

用法（在仓库根目录执行）：
    python examples/make_demo.py
"""
from __future__ import annotations

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

HERE = os.path.dirname(os.path.abspath(__file__))
ROSTER = os.path.join(HERE, "示例学生名册.xlsx")
TEMPLATE = os.path.join(HERE, "示例档案表模板.xlsx")

HEADER = ["序号", "姓名", "学号", "学院", "班级", "是否存档", "备注"]
FONT = Font(name="宋体", size=11)
SIDE = Side(style="thin")
BOX = Border(left=SIDE, right=SIDE, top=SIDE, bottom=SIDE)
CENTER = Alignment(horizontal="center", vertical="center")

# (学院, 培养层次, 班级, 班号, 人数)
CLASSES = [
    ("示例理工学院", "普通本科", "计算机2026-1", "01", 5),
    ("示例理工学院", "普通本科", "计算机2026-2", "02", 4),
    ("示例理工学院", "专升本", "计算机(专升本)2026-1", "03", 3),
    ("示例商学院", "普通本科", "会计(产业学院)2026-1", "04", 6),
    ("示例商学院", "专升本", "会计(专升本)2026-1", "05", 3),
]

FOOTER = [
    "应收        份",
    "实收        份",
    "接收人：",
    "移交人：",
    "核查人：",
    "时间：",
]


def make_roster():
    """名册：第 1 行大标题、第 2 行空、第 3 行才是表头（模拟教务系统导出）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "学生信息"
    ws["A1"] = "示例学生名册（全部为虚构数据，仅用于演示）"
    ws["A1"].font = Font(name="宋体", size=12, bold=True)

    cols = ["学号", "姓名", "培养层次", "上课院系", "班级", "班号"]
    for i, name in enumerate(cols, start=1):
        c = ws.cell(row=3, column=i, value=name)
        c.font = FONT
        c.alignment = CENTER

    row, sid = 4, 20260001
    for college, level, cls, code, n in CLASSES:
        for _ in range(n):
            values = [f"D{sid}", f"示例学生{sid - 20260000:02d}", level, college, cls, code]
            for i, v in enumerate(values, start=1):
                c = ws.cell(row=row, column=i, value=v)
                c.font = FONT
            sid += 1
            row += 1

    for col, w in zip("ABCDEF", [14, 16, 12, 18, 26, 8]):
        ws.column_dimensions[col].width = w
    wb.save(ROSTER)
    return row - 4


def make_template():
    """模板：第 1 行表头，第 2~9 行空白数据区（只带边框），第 10~15 行 D 列为页脚。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "模板"
    for i, h in enumerate(HEADER, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(name="宋体", size=11, bold=True)
        c.alignment = CENTER
        c.border = BOX

    for r in range(2, 10):
        for col in range(1, len(HEADER) + 1):
            c = ws.cell(row=r, column=col)
            c.font = FONT
            c.alignment = CENTER
            c.border = BOX

    for i, text in enumerate(FOOTER):
        c = ws.cell(row=10 + i, column=4, value=text)
        c.font = FONT

    for col, w in zip("ABCDEFG", [8, 12, 14, 20, 24, 12, 12]):
        ws.column_dimensions[col].width = w
    wb.save(TEMPLATE)


if __name__ == "__main__":
    n = make_roster()
    make_template()
    print(f"已生成 {n} 名虚构学生 / {len(CLASSES)} 个班级：")
    print("  " + ROSTER)
    print("  " + TEMPLATE)
