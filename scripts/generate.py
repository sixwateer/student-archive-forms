#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""学生档案纸质表生成器。

从「学生名册」+「Excel 模板」出发，按培养层次 / 学院 / 班级拆分，
每个班级一个 sheet，套用模板表头与页脚，并设置为 A4 可打印版式。

用法示例见 SKILL.md；全部参数见 --help。
"""
from __future__ import annotations

import argparse
import glob
import io
import json
import os
import re
import zipfile
from collections import defaultdict
from copy import copy

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Color, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

INCH_PER_CM = 0.3937
UNIT_CM = 0.185                # 1 个 Excel 列宽单位 ≈ 0.185cm（Calibri 11 的 '0'）
FULL_CM = 11 * 2.54 / 72       # 宋体 11pt 全角字符宽
HALF_CM = FULL_CM / 2
PAD = 1.2                      # 单元格左右留白（列宽单位）
A4_WIDTH_CM = 21.0
DEFAULT_FONT = "宋体"
DEFAULT_SIZE = 11
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")

DEFAULT_COLUMN_MAP = {
    "序号": "__index__",
    "姓名": "姓名",
    "学号": "学号",
    "学院": "上课院系",
    "班级": "班级",
    "是否存档": "__blank__",
    "备注": "__blank__",
}


class Report:
    """同时输出到 stdout 和 UTF-8 日志文件（Windows 控制台常把 UTF-8 二次转码成乱码）。"""

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


def is_fullwidth(ch):
    o = ord(ch)
    return (0x2E80 <= o <= 0x9FFF) or (0x3000 <= o <= 0x303F) or (0xFF00 <= o <= 0xFFEF)


def text_cm(s):
    return sum(FULL_CM if is_fullwidth(c) else HALF_CM for c in str(s))


def need_units(s):
    return text_cm(s) / UNIT_CM + PAD


def safe_sheet_name(name, used):
    n = INVALID_SHEET_CHARS.sub("_", str(name)).strip()[:31] or "Sheet"
    base, k = n, 2
    while n in used:
        suffix = f"~{k}"
        n = base[:31 - len(suffix)] + suffix
        k += 1
    used.add(n)
    return n


def cell_style(cell):
    return {
        "font": copy(cell.font),
        "alignment": copy(cell.alignment),
        "border": copy(cell.border),
        "number_format": cell.number_format,
    }


BLACK = "FF000000"


def _b_side(style):
    """实线边；颜色固定为黑（模板里的 auto 色在部分 Excel/WPS 与打印驱动下会不可见）。"""
    if not style or style == "none":
        return Side()
    return Side(style=style, color=Color(rgb=BLACK))


def thin_box():
    s = _b_side("thin")
    return Border(left=s, right=s, top=s, bottom=s)


def apply_borders(ws, ncols, first_row, last_row, inner="thin", frame="medium"):
    """给矩形区域画实线：内部细黑线 + 外框加粗黑线，保证打印出来看得见。

    inner / frame 取 thin|medium|thick|none；none 表示不画。
    """
    if (not inner or inner == "none") and (not frame or frame == "none"):
        return
    for r in range(first_row, last_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                left=_b_side(frame if c == 1 else inner),
                right=_b_side(frame if c == ncols else inner),
                top=_b_side(frame if r == first_row else inner),
                bottom=_b_side(frame if r == last_row else inner),
            )


def split_footer_rows(raw, min_gap):
    """把表头以下的非空单元格切成「示例数据区」与「页脚区」。

    模板常被填入若干示例行。取最大的行间隔作为分界：间隔 >= min_gap 时，
    其后的连续区块才是页脚；否则全部视为页脚。
    返回 (页脚单元格列表, 被忽略的示例单元格数)。
    """
    rows = sorted({r for _, r, _ in raw})
    if not rows:
        return [], 0
    start = rows[0]
    if len(rows) > 1:
        gaps = [(rows[i + 1] - rows[i], i) for i in range(len(rows) - 1)]
        gap, idx = max(gaps)
        if gap >= min_gap:
            start = rows[idx + 1]
    footer = [(c, r, v) for c, r, v in raw if r >= start]
    return footer, len(raw) - len(footer)


def read_template(path, header_row, data_row, footer_gap):
    """提取模板的表头、页脚（相对偏移）和三类单元格样式。"""
    ws = load_workbook(path).worksheets[0]
    header = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            break
        header.append(str(v).strip())

    hdr_style = cell_style(ws.cell(row=header_row, column=1))
    data_style = cell_style(ws.cell(row=data_row, column=1))
    if not any([data_style["border"].left.style, data_style["border"].right.style,
                data_style["border"].top.style, data_style["border"].bottom.style]):
        data_style["border"] = thin_box()

    raw = []
    for row in ws.iter_rows(min_row=header_row + 1):
        for cell in row:
            if cell.value is not None:
                raw.append((cell.column, cell.row, str(cell.value)))
    footer_raw, ignored = split_footer_rows(raw, footer_gap)
    first = min(r for _, r, _ in footer_raw) if footer_raw else None
    footer = [(col, row - first, val) for col, row, val in footer_raw]
    footer_style = cell_style(ws.cell(row=first, column=footer[0][0])) if footer else data_style
    return {
        "header": header,
        "footer": footer,
        "ignored": ignored,
        "hdr_style": hdr_style,
        "data_style": data_style,
        "footer_style": footer_style,
    }


def apply_style(cell, style, align_center=False):
    cell.font = style["font"] or Font(name=DEFAULT_FONT, size=DEFAULT_SIZE)
    cell.border = style["border"]
    a = copy(style["alignment"])
    if align_center:
        a = Alignment(horizontal="center", vertical="center")
    cell.alignment = a
    cell.number_format = style.get("number_format", "General")


def compute_widths(header, column_values, margin_cm, note_min=13.0):
    """按实际字宽反推列宽，并压进 A4 可打印宽度；备注列尽量留手写空间。

    只放宽「备注」这一列，其余列宽不动；加宽量受 max_total 限制，
    不会触发 fitToWidth 整体缩放，因此其他字段位置与打印行为保持不变。
    F「是否存档」与 G「备注」始终共用同一行线，上下边界齐平。
    """
    n = len(header)
    need = [need_units(h) for h in header]
    for i, vals in enumerate(column_values):
        for v in vals:
            need[i] = max(need[i], need_units(v))
    max_total = (A4_WIDTH_CM - 2 * margin_cm) / UNIT_CM
    total = sum(need)

    # 先收缩可压缩列：备注 / 是否存档 / 序号 / 姓名
    for h in ["备注", "是否存档", "序号", "姓名"]:
        if total <= max_total:
            break
        if h not in header:
            continue
        i = header.index(h)
        floor = need_units(h)
        cut = min(need[i] - floor, total - max_total)
        if cut > 0:
            need[i] -= cut
            total -= cut

    # 有富余时把备注列放宽到 note_min
    if "备注" in header and total <= max_total:
        i = header.index("备注")
        add = min(max(0.0, note_min - need[i]), max_total - total)
        need[i] += add
        total += add
    return [round(v, 2) for v in need], total, max_total


def set_a4(ws, ncols, last_row, margin_cm):
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    m = margin_cm * INCH_PER_CM
    for attr in ["left", "right", "top", "bottom"]:
        setattr(ws.page_margins, attr, m)
    ws.page_margins.header = margin_cm * INCH_PER_CM
    ws.page_margins.footer = margin_cm * INCH_PER_CM
    ws.print_area = f"A1:{get_column_letter(ncols)}{last_row}"
    ws.print_title_rows = "1:1"
    ws.sheet_view.zoomScale = 100


def build_sheet(wb, sheet_name, rows, tpl, widths, margin_cm, used_names, border="thin"):
    ws = wb.create_sheet(safe_sheet_name(sheet_name, used_names))
    header = tpl["header"]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for c, h in enumerate(header, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        apply_style(cell, tpl["hdr_style"], align_center=True)

    for r, row in enumerate(rows, start=2):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            # 前 5 列居中；是否存档/备注留给手写，仅垂直居中
            apply_style(cell, tpl["data_style"], align_center=(c <= 5))

    last_data_row = len(rows) + 1
    max_off = 0
    for col, off, val in tpl["footer"]:
        cell = ws.cell(row=last_data_row + 1 + off, column=col, value=val)
        apply_style(cell, tpl["footer_style"])
        max_off = max(max_off, off)
    last_row = last_data_row + 1 + max_off if tpl["footer"] else last_data_row

    # 只为「表头 + 学生行」画实线（外框加粗）；页脚签名区按需求留白、不画线
    inner = "thin" if border in ("medium", "thick") else border
    apply_borders(ws, len(header), 1, last_data_row, inner=inner, frame=border)
    for r in range(last_data_row + 1, last_row + 1):
        for c in range(1, len(header) + 1):
            ws.cell(row=r, column=c).border = Border()

    set_a4(ws, len(header), last_row, margin_cm)
    return ws


def make_groups(df, mode, prefix):
    """返回 [(文件名, [(班级名, 子表), ...]), ...]；按 层次+学院+班级 分组避免同名班级串组。"""
    groups = defaultdict(list)
    for (lv, college, cls), g in df.groupby(["_level", "_college", "_cls"], sort=False):
        if mode == "college-level":
            fname = f"{prefix}-{college}-{lv}.xlsx"
        elif mode == "college":
            fname = f"{prefix}-{college}.xlsx"
        elif mode == "level":
            fname = f"{prefix}-{lv}.xlsx"
        else:
            fname = f"{prefix}-全部班级.xlsx"
        groups[fname].append((cls, g))
    return list(groups.items())


def parse_level_map(s):
    out = {}
    for part in s.split(","):
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip()] = v.strip()
    return out


def split_values(v):
    """把 'a,b' 或 '["a","b"]' 解析成去重后的字符串列表。"""
    if not v:
        return []
    s = str(v).strip()
    if s.startswith("["):
        try:
            items = json.loads(s)
        except ValueError:
            items = [s]
    else:
        items = s.split(",")
    return list(dict.fromkeys(str(x).strip() for x in items if str(x).strip()))


def year4(v):
    """从 2026级 / 2026-09 / '2026' 等写法里取 4 位年份。"""
    m = re.search(r"(20\d{2})", str(v or ""))
    return m.group(1) if m else ""


def grade_columns(df, spec):
    """确定年级列：显式指定优先，否则按「层次所在年级」口径自动探测。

    「入学年份」不作为年级来源：留级 / 转专业会让同一个班混出多个入学年份，
    用它筛年级会把班级误标到别的届，并在生成时把班里的留级生剔掉。
    """
    if spec:
        return [c.strip() for c in str(spec).split(",")
                if c.strip() and c.strip() in df.columns]
    return [c for c in ("层次所在年级", "当前所在级", "年级", "所在级") if c in df.columns]


def show_filters(args):
    """--list-filters：列出名册里可筛选的取值与人数，不生成任何文件。"""
    df = pd.read_excel(args.source, sheet_name=0, header=args.header_row - 1, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df[args.id_column].notna()].copy()
    lmap = parse_level_map(args.level_map)
    print("名册:", args.source)
    print("总人数:", len(df))

    gcols = grade_columns(df, args.grade_column)
    if gcols:
        gs = defaultdict(int)
        for c in gcols:
            for v in df[c]:
                y = year4(v)
                if y:
                    gs[y] += 1
        print("\n年级列:", "、".join(gcols))
        for k in sorted(gs):
            print("  %s级  %d 人次" % (k, gs[k]))
    else:
        print("\n年级列: 未找到（--grade 将只用于文件命名，不筛数据）")

    def show(title, series):
        print("\n" + title + ":")
        for k, v in series.value_counts().items():
            print("  %s  %d 人" % (k, int(v)))

    if args.level_column in df.columns:
        show("培养层次（--level 用显示名）",
             df[args.level_column].astype(str).str.strip().map(lambda v: lmap.get(v, v)))
    if args.college_column in df.columns:
        show("学院（--college）", df[args.college_column].astype(str).str.strip())
    if args.class_column in df.columns:
        show("班级（--class）", df[args.class_column].astype(str).str.strip())
    print("\n示例: --grade 2025 --level 本科 --college 商学院 --class '会计2025-1'")


def build_parser():
    p = argparse.ArgumentParser(description="学生档案纸质表生成器", formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--config", help="JSON 配置文件，里面同名字段作为默认值；命令行参数可覆盖")
    p.add_argument("--source", default="", help="学生名册 .xls/.xlsx 路径")
    p.add_argument("--template", default="", help="档案纸质表格模板 .xlsx 路径")
    p.add_argument("--out", default="", help="输出目录")
    p.add_argument("--header-row", type=int, default=3, help="名册表头所在行号（1 起）")
    p.add_argument("--template-header-row", type=int, default=1, help="模板表头所在行号")
    p.add_argument("--template-data-row", type=int, default=2, help="模板首个数据行（用于取样式）")
    p.add_argument("--footer-gap", type=int, default=3,
                   help="模板示例数据区与页脚之间的最小空行数；用于自动区分两者")
    p.add_argument("--group-mode", choices=["college-level", "college", "level", "class"], default="college-level")
    p.add_argument("--id-column", default="学号")
    p.add_argument("--name-column", default="姓名")
    p.add_argument("--level-column", default="培养层次")
    p.add_argument("--college-column", default="上课院系")
    p.add_argument("--class-column", default="班级")
    p.add_argument("--class-code-column", default="班号", help="班级排序码；留空则按班级名排")
    p.add_argument("--column-map", default="", help='覆盖表头→名册列映射的 JSON，如 \'{"学院":"所属学院"}\'')
    p.add_argument("--level-map", default="普通本科=本科,专升本=专升本", help='培养层次重命名，逗号分隔')
    p.add_argument("--name-prefix", default="", help="输出文件名前缀；默认用名册文件名")
    p.add_argument("--margin-cm", type=float, default=2.0, help="A4 四周边距（cm）")
    p.add_argument("--border-style", default="thin", dest="border_style",
                   choices=["none", "thin", "medium", "thick"],
                   help="表格实线样式：thin=全细黑线（默认，外框与内线一致）；"
                        "medium/thick=外框加粗；none=沿用模板原有边框")
    p.add_argument("--note-width", type=float, default=13.0,
                   help="备注列目标宽度（列宽单位，默认 13.0 留出更长手写空白）；只影响备注列，有富余时才放宽")
    p.add_argument("--stats", action="store_true", help="输出 学院×层次 人数与占比统计")
    p.add_argument("--zip-by-level", action="store_true", help="按层次把生成的文件打成 zip（配合 college-level 模式）")
    p.add_argument("--exclude", default="", help="逗号分隔的层次显示名，生成时跳过")
    p.add_argument("--grade", default="", dest="grade",
                   help="年级，如 2025 或 2025级；按「当前所在级（层次所在年级）」筛班级，并自动写进文件名")
    p.add_argument("--grade-column", default="", dest="grade_column",
                   help="年级列名，逗号分隔可写多个；留空自动探测 层次所在年级/当前所在级/年级/所在级（不含入学年份）")
    p.add_argument("--level", default="", dest="level_filter",
                   help="只保留的层次显示名，逗号分隔，如 本科 / 专升本 / 本科,专升本；留空=全部")
    p.add_argument("--college", default="", dest="college_filter",
                   help="只保留的学院，逗号分隔或 JSON 数组；留空=全部")
    p.add_argument("--class", default="", dest="class_filter",
                   help="只保留的班级，逗号分隔或 JSON 数组；留空=全部")
    p.add_argument("--exclude-ids", default="", dest="exclude_ids",
                   help="排除的学号，逗号分隔；也可写 @文件路径（每行一个学号）")
    p.add_argument("--list-filters", action="store_true", dest="list_filters",
                   help="只列出名册里可筛选的年级/层次/学院/班级及人数，不生成文件")
    p.add_argument("--log", default="", help="报告文件路径；默认 <out>/_生成报告.txt")
    return p


def main():
    # 先解析 --config，把配置设为 argparse 默认值
    pre = argparse.ArgumentParser(add_help=False)
    pre.add_argument("--config")
    known, _ = pre.parse_known_args()
    parser = build_parser()
    if known.config:
        with io.open(known.config, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        # 配置文件里用更自然的键名 level / college / class，映射到 argparse 的 dest
        alias = {"level": "level_filter", "college": "college_filter",
                 "class": "class_filter"}
        cfg = {alias.get(k, k): v for k, v in cfg.items()}
        valid = {a.dest for a in parser._actions}
        parser.set_defaults(**{k: v for k, v in cfg.items() if k in valid})
    args = parser.parse_args()
    # --list-filters 只需名册，不需要模板与输出目录
    if args.list_filters:
        if not args.source:
            parser.error("缺少 --source（也可在 --config 的 JSON 里提供）")
        show_filters(args)
        return
    for key in ("source", "template", "out"):
        if not getattr(args, key):
            parser.error(f"缺少 --{key}（也可在 --config 的 JSON 里提供）")

    os.makedirs(args.out, exist_ok=True)
    rep = Report(args.log or os.path.join(args.out, "_生成报告.txt"))
    try:
        run(args, rep)
    except Exception as e:
        rep("!!! 失败:", type(e).__name__, str(e))
        raise
    finally:
        rep.close()


def run(args, rep):
    rep("=" * 70)
    rep("学生档案纸质表生成报告")
    rep("=" * 70)
    rep("名册:", args.source)
    rep("模板:", args.template)
    rep("输出:", args.out, "| 分组模式:", args.group_mode)

    df = pd.read_excel(args.source, sheet_name=0, header=args.header_row - 1, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df[args.id_column].notna()].copy()
    for c in set([args.id_column, args.name_column, args.level_column,
                  args.college_column, args.class_column, args.class_code_column]):
        if c and c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    rep("名册列:", list(df.columns))
    rep("学生行数:", len(df), "| 班级数:", df[args.class_column].nunique())

    tpl = read_template(args.template, args.template_header_row,
                        args.template_data_row, args.footer_gap)
    header = tpl["header"]
    rep("模板表头:", header)
    rep("模板页脚单元格数:", len(tpl["footer"]),
        "| 已忽略示例数据单元格:", tpl["ignored"],
        f"（--footer-gap={args.footer_gap}）")
    if len(tpl["footer"]) > 20:
        rep("  ⚠ 页脚单元格偏多，检查 --footer-gap 或清空模板示例数据")

    colmap = dict(DEFAULT_COLUMN_MAP)
    # 名册列改名时同步映射
    colmap["姓名"] = args.name_column
    colmap["学号"] = args.id_column
    colmap["学院"] = args.college_column
    colmap["班级"] = args.class_column
    if args.column_map:
        colmap.update(json.loads(args.column_map))
    missing = [h for h in header if colmap.get(h) not in (None, "__index__", "__blank__")
               and colmap[h] not in df.columns]
    if missing:
        raise ValueError(f"表头找不到对应名册列: {missing}；用 --column-map 修正。可选列: {list(df.columns)}")

    lmap = parse_level_map(args.level_map)
    level_order = list(dict.fromkeys(lmap.values())) or sorted(df[args.level_column].unique())
    df["_level"] = df[args.level_column].map(lambda v: lmap.get(str(v).strip(), str(v).strip()))
    df["_college"] = df[args.college_column].astype(str).str.strip()
    df["_cls"] = df[args.class_column].astype(str).str.strip()
    df["_lv_order"] = df["_level"].map({lv: i for i, lv in enumerate(level_order)}).fillna(99)
    df["_code"] = df[args.class_code_column].astype(str) if args.class_code_column in df.columns else df["_cls"]
    df["_id"] = df[args.id_column].astype(str)
    df = df.sort_values(by=["_lv_order", "_college", "_code", "_id", "_cls"], kind="stable")

    rep("层次取值:", {k: int(v) for k, v in df["_level"].value_counts().items()})

    # ---- 先筛选，再生成 ----
    conditions = []
    g4 = year4(args.grade)
    gcols = grade_columns(df, args.grade_column)
    if g4:
        conditions.append("年级=%s级" % g4)

    lv_keep = split_values(args.level_filter)
    if lv_keep and not any(x in ("全部", "all", "ALL") for x in lv_keep):
        conditions.append("层次=" + "/".join(lv_keep))
        before = len(df)
        raw_lv = df[args.level_column].astype(str).str.strip()
        df = df[df["_level"].isin(lv_keep) | raw_lv.isin(lv_keep)].copy()
        rep("按层次筛选:", lv_keep, f"| 剔除 {before-len(df)} 行")

    col_keep = split_values(args.college_filter)
    if col_keep:
        conditions.append("学院=" + "/".join(col_keep))
        before = len(df)
        df = df[df["_college"].isin(col_keep)].copy()
        rep("按学院筛选:", col_keep, f"| 剔除 {before-len(df)} 行")

    cls_keep = split_values(args.class_filter)
    if cls_keep:
        conditions.append("班级=" + "/".join(cls_keep))
        before = len(df)
        df = df[df["_cls"].isin(cls_keep)].copy()
        rep("按班级筛选:", cls_keep, f"| 剔除 {before-len(df)} 行")

    # 年级只用于决定「哪些班级入选」与文件命名。按班级字段筛选后，该班在原表格里的
    # 全部学生（含留级 / 转专业）必须一起进表，不得再按年级把同一个班拆开。
    if g4 and not cls_keep:
        if gcols:
            mask = pd.Series(False, index=df.index)
            for c in gcols:
                mask |= df[c].map(year4).eq(g4)
            before = len(df)
            df = df[mask].copy()
            rep("按年级筛选:", g4, "级 | 年级列:", gcols, f"| 剔除 {before-len(df)} 行")
        else:
            rep("按年级筛选:", g4, "级 | ⚠ 名册无年级列，只用于命名，不筛数据")

    ex_ids = split_values(args.exclude_ids)
    if str(args.exclude_ids).strip().startswith("@"):
        fp = str(args.exclude_ids).strip()[1:]
        with io.open(fp, "r", encoding="utf-8-sig") as f:
            ex_ids = [x.strip() for x in f.read().splitlines() if x.strip()]
    if ex_ids:
        conditions.append("排除学号=%d个" % len(ex_ids))
        before = len(df)
        df = df[~df["_id"].isin(set(ex_ids))].copy()
        rep("排除学号:", len(ex_ids), f"个 | 剔除 {before-len(df)} 行")

    if conditions:
        rep("筛选条件:", "；".join(conditions))
        rep("筛选后:", len(df), "人 |", df["_cls"].nunique(), "个班")
    if df.empty:
        raise ValueError("筛选后没有学生，请放宽 --grade / --level / --college / --class 条件。")

    # 把筛选结果落盘，供 verify.py 自动按同一批学生核对（放在子目录，避免被当成成果文件）
    meta = {
        "enabled": bool(conditions),
        "conditions": conditions,
        "source": args.source,
        "header_row": 1,
        "subset": "",
        "created_by": "generate.py",
    }
    if conditions:
        sub_dir = os.path.join(args.out, "_筛选信息")
        os.makedirs(sub_dir, exist_ok=True)
        sub_path = os.path.join(sub_dir, "筛选后名册.xlsx")
        df.drop(columns=[c for c in df.columns if c.startswith("_")]).to_excel(sub_path, index=False)
        meta["subset"] = sub_path
        rep("已保存筛选后名册:", sub_path, "（校验会自动使用它）")
    with io.open(os.path.join(args.out, "_筛选条件.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    if args.exclude:
        skip = set(x.strip() for x in args.exclude.split(","))
        before = len(df)
        df = df[~df["_level"].isin(skip)].copy()
        rep("按 --exclude 排除:", sorted(skip), f"剔除 {before-len(df)} 行")

    # 列宽：收集所有将被写入的字符串
    colvals = [set() for _ in header]
    for i, h in enumerate(header):
        src = colmap.get(h)
        if src and src not in ("__index__", "__blank__"):
            colvals[i].update(df[src].dropna().unique().tolist())
        elif src == "__index__":
            colvals[i].add(str(len(df)))
    for col, off, val in tpl["footer"]:
        if col <= len(header):
            colvals[col - 1].add(val)
    widths, total_units, max_units = compute_widths(header, colvals, args.margin_cm, args.note_width)
    rep("列宽:", dict(zip([get_column_letter(i + 1) for i in range(len(header))], widths)))
    rep("列宽合计 %.1f 单位 ≈ %.1fcm / 上限 %.1fcm -> %s" %
        (total_units, total_units * UNIT_CM, max_units * UNIT_CM,
         "100%比例可容纳" if total_units <= max_units else "将缩印"))

    prefix = (args.name_prefix or (f"{g4}级档案纸质表格" if g4 else "")
              or os.path.splitext(os.path.basename(args.source))[0])
    groups = make_groups(df, args.group_mode, prefix)

    made = []
    total_sheets = total_rows = 0
    for fname, classes in groups:
        # 文件名去掉非法字符
        fname = INVALID_SHEET_CHARS.sub("_", fname).replace("/", "_").replace("\\", "_")
        if any(ch in fname for ch in ':*?"<>|'):
            for ch in ':*?"<>|':
                fname = fname.replace(ch, "_")
        wb = Workbook()
        wb.remove(wb.active)
        used = set()
        for cls, g in classes:
            rows = []
            for i, rec in enumerate(g.to_dict("records"), start=1):
                row = []
                for h in header:
                    src = colmap.get(h)
                    if src == "__index__":
                        row.append(i)
                    elif src in (None, "__blank__"):
                        row.append(None)
                    else:
                        row.append(rec.get(src))
                rows.append(row)
            build_sheet(wb, cls, rows, tpl, widths, args.margin_cm, used,
                        border=args.border_style)
            total_rows += len(rows)
            total_sheets += 1
        path = os.path.join(args.out, fname)
        wb.save(path)
        made.append(path)
        rep(f"  {fname}  sheet={len(classes)}  人={sum(len(g) for _, g in classes)}")

    rep("-" * 70)
    rep("合计: 文件=%d sheet=%d 学生行=%d | 源表 %d 人 / %d 班" %
        (len(made), total_sheets, total_rows, len(df), df["_cls"].nunique()))

    if args.stats:
        rep("\n### 学院 × 层次 人数统计")
        ct = pd.crosstab(df["_college"], df["_level"])
        # 筛选/排除后可能只剩一种层次，把缺失的层次补成 0 列，避免 KeyError
        for lv in level_order:
            if lv not in ct.columns:
                ct[lv] = 0
        ct["合计"] = ct[level_order].sum(axis=1)
        ct = ct.sort_values("合计", ascending=False)
        grand = ct["合计"].sum()
        rep(f"{'学院':<14}" + "".join(f"{lv:>10}" for lv in level_order) + f"{'合计':>10}{'占全校':>9}")
        for idx, r in ct.iterrows():
            rep(f"{idx:<14}" + "".join(f"{int(r[lv]):>10}" for lv in level_order) +
                f"{int(r['合计']):>10}{(r['合计']/grand*100 if grand else 0):>8.1f}%")
        rep(f"{'全校':<14}" + "".join(f"{int(ct[lv].sum()):>10}" for lv in level_order) +
            f"{int(grand):>10}{(100.0 if grand else 0):>8.1f}%")

    if args.zip_by_level:
        for lv in level_order:
            fs = sorted(f for f in made if os.path.basename(f).endswith(f"-{lv}.xlsx"))
            if not fs:
                continue
            zpath = os.path.join(args.out, f"{prefix}-{lv}({len(fs)}个文件).zip")
            with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
                for f in fs:
                    z.write(f, os.path.basename(f))
            rep(f"  打包 {os.path.basename(zpath)}  含 {len(fs)} 个文件  {os.path.getsize(zpath)/1024:.0f} KB")

    rep("\n下一步: 运行 verify.py 校验输出（人数、学号、打印设置）。")


if __name__ == "__main__":
    main()
