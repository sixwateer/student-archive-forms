# -*- coding: utf-8 -*-
"""
合并散落的学生档案成果，打包成一个大 ZIP，方便整体发给别人。

场景：生成时按学院 / 层次拆成了很多个 .xlsx（或若干个「本科(9学院).zip」「专升本(8学院).zip」），
      现在要把它们合成「一个」大 ZIP。

输入可以是任意混合：
  · 文件夹            —— 收集里面的 *.xlsx（含可选 *_生成报告.txt / *_校验报告.txt）
  · .zip 文件         —— 把里面的 *.xlsx（及报告）展平收进来
  · .xlsx / .txt 文件 —— 直接收进来

同名文件自动去重（冲突时加「来源__」前缀）；写入后校验完整性，并附 _合并清单.txt。
"""
from __future__ import annotations

import argparse
import os
import sys
import zipfile
from datetime import datetime

# 这些临时 / 系统文件不收进大 ZIP
SKIP_NAMES = {"_服务地址.txt", "_合并清单.txt"}
XLSX_SUFFIX = ".xlsx"


def _good_report(name):
    base = os.path.basename(name)
    if base in SKIP_NAMES:
        return False
    # 其余以 _ 开头的报告（_生成报告.txt / _校验报告.txt）保留
    return True


def _iter_sources(inputs, include_history, include_reports):
    """逐个产出 (source_label, source_desc, entry_name, data_bytes, is_report)。"""
    for inp in inputs:
        if not os.path.exists(inp):
            print("  ⚠ 跳过不存在的路径：%s" % inp, file=sys.stderr)
            continue

        if os.path.isfile(inp):
            low = inp.lower()
            if low.endswith(".zip"):
                yield from _iter_zip(inp, include_reports)
            elif low.endswith(XLSX_SUFFIX):
                with open(inp, "rb") as f:
                    data = f.read()
                yield (os.path.splitext(os.path.basename(inp))[0],
                       os.path.basename(inp), os.path.basename(inp), data, False)
            elif low.endswith(".txt") and include_reports and _good_report(inp):
                with open(inp, "rb") as f:
                    data = f.read()
                yield (os.path.splitext(os.path.basename(inp))[0],
                       os.path.basename(inp), os.path.basename(inp), data, True)
            continue

        # 目录
        label = os.path.basename(os.path.normpath(inp)) or inp
        for fn in sorted(os.listdir(inp)):
            p = os.path.join(inp, fn)
            if not os.path.isfile(p):
                continue
            low = p.lower()
            if low.endswith(XLSX_SUFFIX):
                with open(p, "rb") as f:
                    data = f.read()
                yield (label, fn, fn, data, False)
            elif low.endswith(".txt") and include_reports and _good_report(p):
                with open(p, "rb") as f:
                    data = f.read()
                yield (label, fn, fn, data, True)

        if include_history:
            hist = os.path.join(inp, "_历史版本")
            if os.path.isdir(hist):
                for sub in sorted(os.listdir(hist)):
                    subp = os.path.join(hist, sub)
                    if not os.path.isdir(subp):
                        continue
                    sub_label = "%s/_历史版本/%s" % (label, sub)
                    for fn in sorted(os.listdir(subp)):
                        p = os.path.join(subp, fn)
                        if not os.path.isfile(p):
                            continue
                        low = p.lower()
                        if low.endswith(XLSX_SUFFIX):
                            with open(p, "rb") as f:
                                data = f.read()
                            yield (sub_label, fn, fn, data, False)
                        elif low.endswith(".txt") and include_reports and _good_report(p):
                            with open(p, "rb") as f:
                                data = f.read()
                            yield (sub_label, fn, fn, data, True)


def _iter_zip(zippath, include_reports):
    label = os.path.splitext(os.path.basename(zippath))[0]
    with zipfile.ZipFile(zippath) as z:
        for n in z.namelist():
            if n.endswith("/"):
                continue
            low = n.lower()
            is_report = low.endswith(".txt")
            if low.endswith(XLSX_SUFFIX):
                yield (label, "%s › %s" % (os.path.basename(zippath), n), n, z.read(n), False)
            elif is_report and include_reports and _good_report(n):
                yield (label, "%s › %s" % (os.path.basename(zippath), n), n, z.read(n), True)


def combine(inputs, out=None, include_history=False, include_reports=True, log=None):
    """把 inputs 合并成一个大 ZIP。返回结果字典。"""
    inputs = [os.path.abspath(p) for p in inputs]
    # 默认输出到第一个【文件夹】输入同级，或第一个输入所在目录
    if not out:
        base_dir = None
        for p in inputs:
            if os.path.isdir(p):
                base_dir = p
                break
        if not base_dir:
            base_dir = os.path.dirname(inputs[0]) if inputs else "."
        out = os.path.join(base_dir, "档案表成果-合并全部.zip")

    out = os.path.abspath(out)
    used = {}
    manifest_rows = []
    n_xlsx = n_rep = 0
    collisions = 0
    total_bytes = 0

    # 先把待写入项收集好，统一处理去重
    pending = []  # (final_name, source_desc, data, is_report)
    for source_label, source_desc, entry_name, data, is_report in _iter_sources(
            inputs, include_history, include_reports):
        base = os.path.basename(entry_name)
        final = base
        if final in used:
            collisions += 1
            final = "%s__%s" % (source_label, base)
            # 极端情况仍冲突，再加序号
            i = 2
            while final in used:
                final = "%s__%s(%d)" % (source_label, base, i)
                i += 1
        used[final] = True
        pending.append((final, source_desc, data, is_report))

    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for final, source_desc, data, is_report in pending:
            z.writestr(final, data)
            total_bytes += len(data)
            if is_report:
                n_rep += 1
            else:
                n_xlsx += 1
            manifest_rows.append((final, source_desc, len(data)))

        # 写合并清单
        lines = ["档案表成果合并清单", "=" * 40,
                 "合并时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 "来源数量：%d" % len(inputs),
                 "表格文件：%d 个" % n_xlsx,
                 "报告文件：%d 个" % n_rep,
                 "重名已处理：%d 处" % collisions,
                 "合计大小：%.2f MB" % (total_bytes / 1024 / 1024),
                 "-" * 40,
                 "文件名 | 来源 | 大小"]
        for final, source_desc, size in manifest_rows:
            lines.append("%s | %s | %d 字节" % (final, source_desc, size))
        z.writestr("_合并清单.txt", "\n".join(lines) + "\n")

    # 校验
    bad = zipfile.ZipFile(out).testzip()

    info = {
        "ok": bad is None,
        "out": out,
        "n_xlsx": n_xlsx,
        "n_reports": n_rep,
        "collisions": collisions,
        "total_bytes": total_bytes,
        "size_mb": total_bytes / 1024 / 1024,
        "bad": bad,
    }

    report = [
        "合并完成" if info["ok"] else "合并完成（但校验发现损坏：%s）" % bad,
        "输出：%s" % out,
        "表格文件：%d 个" % n_xlsx,
        "报告文件：%d 个" % n_rep,
        "重名已处理：%d 处" % collisions,
        "大 ZIP 大小：%.2f MB" % info["size_mb"],
        "完整性校验：%s" % ("通过" if info["ok"] else "失败 (%s)" % bad),
    ]
    text = "\n".join(report) + "\n"
    print(text)
    if log:
        try:
            with open(log, "w", encoding="utf-8") as f:
                f.write(text)
        except Exception:
            pass
    return info


def main():
    ap = argparse.ArgumentParser(description="合并散落的学生档案成果成一个大 ZIP")
    ap.add_argument("--in", action="append", required=True, dest="inputs", metavar="路径",
                    help="输入（可多次）：文件夹 / .zip / .xlsx / .txt，混合均可")
    ap.add_argument("--out", default="", help="输出大 ZIP 路径（默认：首个文件夹同级/档案表成果-合并全部.zip）")
    ap.add_argument("--include-history", action="store_true",
                    help="同时把每个文件夹下的「_历史版本」也合并进来")
    ap.add_argument("--no-reports", action="store_true",
                    help="不收报告类 .txt 文件，只收表格 .xlsx")
    ap.add_argument("--log", default="", help="把报告同时写到这个文件")
    args = ap.parse_args()

    info = combine(args.inputs, args.out or None,
                   include_history=args.include_history,
                   include_reports=not args.no_reports,
                   log=args.log or None)
    sys.exit(0 if info["ok"] else 1)


if __name__ == "__main__":
    main()
