# -*- coding: utf-8 -*-
"""
合同数据统计核对 Skill —— 主程序。

一条命令完成：读取源表 → 数据清洗合并 → 未回款原因分类汇总 → 多维度统计
→ 全工作表统一美化 → 公式缓存值注入 → 输出 8 张工作表的结果 Excel。

用法：
    python3 analyze.py --input 附件1-合同开票及回款核对表.xlsx \
                       --output 合同开票及回款核对表分析结果_张三.xlsx

    可选： --profile spec|example   口径/样式规范来源，默认 spec（题目正文）
           --raw-strict             「原始数据」表完全不加任何美化（最严格的原样保留）
           --soffice-recalc         额外用 LibreOffice 重算一次做交叉验证（非必需）
           --report report.json     把执行回执写成 JSON

设计约束（对应题目三大硬性规则）：
    1) 所有统计单元格写入动态公式，绝不写 Python 算出的常量；
       Python 计算结果只用于"公式缓存值注入"和独立复核，两条路径互不复用。
    2) 「原始数据」表的值、行列顺序、数字格式一律不动。
    3) 全部工作表统一美化，参数集中在 config.py / styling.py。
"""

import argparse
import collections
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import cached_values as CV
import config as C
import formulas as F
import styling as S


# ================================================================ 读取源数据

def detect_layout(ws):
    """动态探测表头行与数据区间，不写死行号（题目只保证表头在第 2 行）。"""
    header_row = None
    for r in range(1, min(ws.max_row, 20) + 1):
        values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if "合同号" in [str(v).strip() if v is not None else "" for v in values]:
            header_row = r
            break
    if header_row is None:
        raise ValueError("未在前 20 行找到包含「合同号」的表头行，请确认输入文件")

    col_of = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is None:
            continue
        key = str(v).strip()
        if key and key not in col_of:
            col_of[key] = c

    missing = [f for f in C.TASK_FIELDS if f != "序号" and f not in col_of]
    if missing:
        raise ValueError("源文件缺少必需字段：%s" % "、".join(missing))

    last_row = header_row
    for r in range(header_row + 1, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
            last_row = r
    return header_row, last_row, col_of


def read_records(ws, header_row, last_row, col_of):
    """读全量明细。数值字段的空值按 0 计入并单独计数，不静默丢弃。"""
    records = []
    blank_numeric = collections.Counter()
    for r in range(header_row + 1, last_row + 1):
        rec = {"_row": r}
        for field, c in col_of.items():
            rec[field] = ws.cell(r, c).value
        for field in C.MERGE_SUM_FIELDS:
            v = rec.get(field)
            if v is None or (isinstance(v, str) and not v.strip()):
                blank_numeric[field] += 1
                rec[field] = 0.0
            else:
                rec[field] = float(v)
        records.append(rec)
    return records, dict(blank_numeric)


def is_blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


def text_of(v):
    """保留源文本原样（含前导零），仅去首尾空白用于比较键。"""
    return v if isinstance(v, str) else ("" if v is None else str(v))


# ================================================================ 业务计算（独立于公式）

def compute(records, col_of, profile):
    """用纯 Python 独立算出全部期望值，供缓存注入与复核使用。"""
    reason_field = "未回款原因分类"
    res = {}

    res["total_rows"] = len(records)
    res["totals"] = {f: sum(r[f] for r in records) for f in C.MERGE_SUM_FIELDS}
    res["total_invoice"] = res["totals"]["开票金额"]
    res["total_amount"] = res["totals"]["合同金额"]
    res["total_received"] = res["totals"]["回款合计"]
    res["total_unpaid"] = res["totals"]["开票未回款"]

    # ---- 任务1：按合同号分组
    groups = collections.OrderedDict()
    for rec in records:
        code = text_of(rec.get("合同号")).strip()
        g = groups.get(code)
        if g is None:
            g = groups[code] = {
                "code": code,
                "first_row": rec["_row"],
                "sums": {f: 0.0 for f in C.MERGE_SUM_FIELDS},
                "first": {f: None for f in C.MERGE_FIRST_FIELDS},
                "reasons": [],
                "row_count": 0,
            }
        g["row_count"] += 1
        for f in C.MERGE_SUM_FIELDS:
            g["sums"][f] += rec[f]
        for f in C.MERGE_FIRST_FIELDS:
            if is_blank(g["first"][f]) and not is_blank(rec.get(f)):
                g["first"][f] = rec[f]
        rv = rec.get(reason_field)
        if profile["fill_blank_reason"] and is_blank(rv):
            rv = "未填写"
        if not is_blank(rv):
            t = text_of(rv).strip()
            if t not in g["reasons"]:
                g["reasons"].append(t)
    for g in groups.values():
        g["reason_text"] = C.JOIN_SEP.join(g["reasons"])
    res["contract_count"] = len(groups)
    # 排序：附件2-示例 的 463 行逐行验证为「合同金额降序」
    res["contracts"] = sorted(groups.values(), key=lambda g: (-g["sums"]["合同金额"], g["first_row"]))

    # ---- 任务3：四个维度
    dims = {}
    for _sheet, dim in C.DIMENSIONS:
        bucket = collections.OrderedDict()
        for rec in records:
            if profile["fill_blank_reason"] and dim == reason_field and is_blank(rec.get(dim)):
                key = "未填写"
            else:
                if is_blank(rec.get(dim)):
                    continue
                key = text_of(rec[dim]).strip()
            b = bucket.get(key)
            if b is None:
                b = bucket[key] = {
                    "key": key,
                    "sums": {f: 0.0 for f in C.MERGE_SUM_FIELDS},
                    "rows": 0,
                    "codes": set(),
                }
            b["rows"] += 1
            b["codes"].add(text_of(rec.get("合同号")).strip())
            for f in C.MERGE_SUM_FIELDS:
                b["sums"][f] += rec[f]
        ordered = sorted(bucket.values(), key=lambda b: (-b["sums"]["开票未回款"], b["key"]))
        for b in ordered:
            amt = b["sums"]["合同金额"]
            b["rate"] = (b["sums"]["回款合计"] / amt) if amt else None
            b["unpaid_ratio"] = (b["sums"]["开票未回款"] / res["total_amount"]) if res["total_amount"] else None
        dims[dim] = ordered
    res["dims"] = dims

    # ---- 任务2：未回款原因分类
    bucket = collections.OrderedDict()
    for rec in records:
        rv = rec.get(reason_field)
        if profile["fill_blank_reason"] and is_blank(rv):
            rv = "未填写"
        if is_blank(rv):
            continue
        key = text_of(rv).strip()
        b = bucket.get(key)
        if b is None:
            b = bucket[key] = {"key": key, "unpaid": 0.0, "rows": 0, "codes": set()}
        b["unpaid"] += rec["开票未回款"]
        b["rows"] += 1
        b["codes"].add(text_of(rec.get("合同号")).strip())
    reasons = sorted(bucket.values(), key=lambda b: (-b["unpaid"], b["key"]))
    for b in reasons:
        b["ratio"] = (b["unpaid"] / res["total_unpaid"]) if res["total_unpaid"] else None
    res["reasons"] = reasons

    # ---- 统计总览
    res["overview"] = [
        res["total_amount"], res["total_invoice"], res["total_received"], res["total_unpaid"],
        (res["total_received"] / res["total_amount"]) if res["total_amount"] else None,
        res["contract_count"], res["total_rows"], len(reasons),
    ]
    return res


# ================================================================ 写表

class SheetBuilder(object):
    """统一处理：写公式 + 记录期望缓存值 + 记录列宽用的显示值。"""

    def __init__(self, ws, ncols, profile):
        self.ws = ws
        self.ncols = ncols
        self.profile = profile
        self.money_format = profile["money_format"]
        self.expected = {}
        self.display = []
        self.row_kinds = []

    def add_row(self, row_idx, values, kinds):
        """values: 每列 (formula_or_value, expected_value) 二元组"""
        disp = []
        for i, (val, exp) in enumerate(values):
            col = i + 1
            cell = self.ws.cell(row_idx, col)
            cell.value = val
            coord = cell.coordinate
            if exp is not None:
                self.expected[coord] = exp
            disp.append(exp if exp is not None else val)
        self.display.append(disp)
        self.row_kinds.append(kinds)


def build_raw_sheet(ws, profile, header_row, last_row, ncols, col_of):
    """「原始数据」：值/行列顺序/数字格式一律不动，只按规范叠加美化。"""
    kinds = []
    for c in range(1, ncols + 1):
        name = None
        for k, v in col_of.items():
            if v == c:
                name = k
                break
        kinds.append(C.FIELD_KIND.get(name, "text"))
    align = profile["align_by_kind"]

    if profile["raw_override_title_row"]:
        for c in range(1, ncols + 1):
            ws.cell(1, c).value = None
        S.write_title(ws, ncols, C.SHEET_TITLES[C.SHEET_RAW])
    S.style_header(ws, ncols, header_row=header_row)
    S.header_border(ws, ncols, profile["border_mode"], header_row=header_row)

    display = []
    for r in range(header_row + 1, last_row + 1):
        row = []
        for c in range(1, ncols + 1):
            v = ws.cell(r, c).value
            row.append(v)
        display.append(row)

    S.style_data(ws, ncols, header_row + 1, last_row - header_row, kinds, align,
                 profile["border_mode"], set_number_format=False)
    return kinds, display


def build_merge_sheet(ws, res, profile, raw_letters, first_row=3):
    cols = profile["merge_col_order"]
    headers = list(cols)
    kinds = [C.FIELD_KIND.get(f, "text") for f in cols]
    letter_of = {f: get_column_letter(i + 1) for i, f in enumerate(cols)}
    code_letter = letter_of["合同号"]
    raw_code = raw_letters["合同号"]

    b = SheetBuilder(ws, len(cols), profile)
    for i, g in enumerate(res["contracts"]):
        r = first_row + i
        key_ref = "$%s%d" % (code_letter, r)
        values = []
        for f in cols:
            if f == "序号":
                values.append((F.seq_by_row(first_row - 1), i + 1))
            elif f == "合同号":
                values.append((g["code"], g["code"]))
            elif f in C.MERGE_SUM_FIELDS:
                values.append((F.sumif(raw_code, key_ref, raw_letters[f],
                                       C.SRC_DATA_FIRST_ROW, res["_raw_last"]),
                               g["sums"][f]))
            elif f in C.MERGE_FIRST_FIELDS:
                exp = g["first"][f]
                exp = "" if exp is None else exp
                values.append((F.first_value(raw_code, key_ref, raw_letters[f],
                                             C.SRC_DATA_FIRST_ROW, res["_raw_last"]), exp))
            elif f in C.MERGE_JOIN_FIELDS:
                # 去重拼接是文本聚合，Excel 无稳定的非数组公式写法；
                # 此处写入计算结果文本，并在构建过程文档中登记为公式化例外。
                values.append((g["reason_text"], g["reason_text"]))
            else:
                values.append((None, None))
        b.add_row(r, values, kinds)
    return b, headers, kinds


def build_dim_sheet(ws, dim, res, profile, raw_letters, first_row=3):
    headers = [h.replace("{dim}", dim) for h in profile["dim_headers"]]
    kinds = ["int", "cat", "money", "money", "money", "money", "int", "pct", "pct"]
    raw_dim = raw_letters[dim]
    raw_code = raw_letters["合同号"]
    lo, hi = C.SRC_DATA_FIRST_ROW, res["_raw_last"]

    b = SheetBuilder(ws, len(headers), profile)
    for i, item in enumerate(res["dims"][dim]):
        r = first_row + i
        key_ref = "$B%d" % r
        values = [
            (F.seq_by_row(first_row - 1), i + 1),
            (item["key"], item["key"]),
            (F.sumif(raw_dim, key_ref, raw_letters["合同金额"], lo, hi), item["sums"]["合同金额"]),
            (F.sumif(raw_dim, key_ref, raw_letters["开票金额"], lo, hi), item["sums"]["开票金额"]),
            (F.sumif(raw_dim, key_ref, raw_letters["回款合计"], lo, hi), item["sums"]["回款合计"]),
            (F.sumif(raw_dim, key_ref, raw_letters["开票未回款"], lo, hi), item["sums"]["开票未回款"]),
            (F.distinct_count_dim(raw_dim, key_ref, raw_code, lo, hi), len(item["codes"])),
            (F.ratio("E%d" % r, "C%d" % r), item["rate"]),
            ("=IF(%s=0,\"\",F%d/%s)" % (F.sum_all(raw_letters["合同金额"], lo, hi)[1:], r,
                                        F.sum_all(raw_letters["合同金额"], lo, hi)[1:]),
             item["unpaid_ratio"]),
        ]
        b.add_row(r, values, kinds)
    return b, headers, kinds


def build_reason_sheet(ws, res, profile, raw_letters, first_row=3):
    headers = list(C.REASON_HEADERS)
    kinds = ["int", "cat", "money", "int", "int", "pct"]
    raw_reason = raw_letters["未回款原因分类"]
    raw_code = raw_letters["合同号"]
    lo, hi = C.SRC_DATA_FIRST_ROW, res["_raw_last"]
    total_unpaid_sql = F.sum_all(raw_letters["开票未回款"], lo, hi)[1:]

    b = SheetBuilder(ws, len(headers), profile)
    for i, item in enumerate(res["reasons"]):
        r = first_row + i
        key_ref = "$B%d" % r
        values = [
            (F.seq_by_row(first_row - 1), i + 1),
            (item["key"], item["key"]),
            (F.sumif(raw_reason, key_ref, raw_letters["开票未回款"], lo, hi), item["unpaid"]),
            (F.distinct_count_dim(raw_reason, key_ref, raw_code, lo, hi), len(item["codes"])),
            (F.countif(raw_reason, key_ref, lo, hi), item["rows"]),
            ("=IF(%s=0,\"\",C%d/%s)" % (total_unpaid_sql, r, total_unpaid_sql), item["ratio"]),
        ]
        b.add_row(r, values, kinds)
    return b, headers, kinds


def build_overview_sheet(ws, res, profile, raw_letters, first_row=3):
    headers = ["指标", "数值"]
    lo, hi = C.SRC_DATA_FIRST_ROW, res["_raw_last"]
    kinds_by_row = [
        ["text", "money"], ["text", "money"], ["text", "money"], ["text", "money"],
        ["text", "pct"], ["text", "int"], ["text", "int"], ["text", "int"],
    ]
    formulas = [
        F.sum_all(raw_letters["合同金额"], lo, hi),
        F.sum_all(raw_letters["开票金额"], lo, hi),
        F.sum_all(raw_letters["回款合计"], lo, hi),
        F.sum_all(raw_letters["开票未回款"], lo, hi),
        '=IF(B%d=0,"",B%d/B%d)' % (first_row, first_row + 2, first_row),
        F.distinct_count_all(raw_letters["合同号"], lo, hi),
        F.count_all(raw_letters["合同号"], lo, hi),
        F.distinct_count_nonblank_all(raw_letters["未回款原因分类"], lo, hi),
    ]
    b = SheetBuilder(ws, 2, profile)
    for i, label in enumerate(C.OVERVIEW_LABELS):
        r = first_row + i
        exp = res["overview"][i]
        b.add_row(r, [(label, label), (formulas[i], exp)], kinds_by_row[i])
    return b, headers, ["text", "num"], kinds_by_row


# ================================================================ 主流程

def finalize(ws, builder, headers, kinds, profile, title, ncols=None,
             row_kinds=None, extra_center=None):
    """写标题/表头 + 统一美化 + 冻结。"""
    ncols = ncols or len(headers)
    if title:
        S.write_title(ws, ncols, title)
    for i, h in enumerate(headers):
        ws.cell(2, i + 1).value = h
    S.style_header(ws, ncols)
    S.header_border(ws, ncols, profile["border_mode"])

    money_format = profile["money_format"]
    display = [[S.format_display(v, (row_kinds[i][j] if row_kinds else kinds[j]), money_format)
                for j, v in enumerate(row)] for i, row in enumerate(builder.display)]
    for c in range(1, ncols + 1):
        ws.cell(2, c).value = headers[c - 1]

    if row_kinds:
        for i, ks in enumerate(row_kinds):
            S.style_data(ws, ncols, 3 + i, 1, ks, profile["align_by_kind"],
                         profile["border_mode"], number_formats={"money": money_format},
                         zebra=False)
        # 斑马纹与边框整区统一刷一遍
        for i in range(len(row_kinds)):
            r = 3 + i
            striped = (i % 2 == 0)
            for c in range(1, ncols + 1):
                cell = ws.cell(r, c)
                cell.fill = S._fill(C.ZEBRA_FILL) if striped else S.PatternFill(fill_type=None)
    else:
        S.style_data(ws, ncols, 3, len(builder.display), kinds, profile["align_by_kind"],
                     profile["border_mode"], number_formats={"money": money_format})

    S.autofit(ws, ncols, headers, display, ["text"] * ncols, money_format)
    S.freeze(ws)


def run(args):
    t0 = time.time()
    profile = C.get_profile(args.profile)
    if args.raw_strict:
        profile["raw_preserve_strict"] = True

    src = os.path.abspath(args.input)
    if not os.path.exists(src):
        raise SystemExit("输入文件不存在：%s" % src)
    out = os.path.abspath(args.output)
    if args.name:
        out = os.path.join(os.path.dirname(out),
                           "合同开票及回款核对表分析结果_%s.xlsx" % args.name)
    if os.path.abspath(out) == src:
        raise SystemExit("输出路径不能与输入文件相同（禁止篡改原始明细数据）")

    # ---- 读源
    wb = load_workbook(src)
    src_ws = None
    for ws in wb.worksheets:
        if ws.sheet_state == "visible":
            src_ws = ws
            break
    if src_ws is None:
        raise SystemExit("输入文件中没有可见工作表")
    hidden = [w.title for w in wb.worksheets if w.sheet_state != "visible"]
    src_name = src_ws.title
    header_row, last_row, col_of = detect_layout(src_ws)
    records, blank_numeric = read_records(src_ws, header_row, last_row, col_of)
    raw_letters = {f: get_column_letter(c) for f, c in col_of.items()}
    ncols_raw = max(col_of.values())

    res = compute(records, col_of, profile)
    res["_raw_last"] = last_row
    res["_raw_first"] = header_row + 1

    # ---- 「原始数据」：直接在源工作簿上改名，保证逐格原样
    src_ws.title = C.SHEET_RAW
    ws_raw = wb[C.SHEET_RAW]

    report = {
        "input": src, "output": out, "profile": profile["name"],
        "source_sheet": src_name, "hidden_sheets": hidden,
        "header_row": header_row, "data_first_row": header_row + 1, "data_last_row": last_row,
        "source_columns": ncols_raw, "records": len(records),
        "blank_numeric_cells": blank_numeric,
        "contract_count": res["contract_count"],
        "reason_category_count": len(res["reasons"]),
        "dim_category_counts": {d: len(res["dims"][d]) for _s, d in C.DIMENSIONS},
    }

    if profile["raw_preserve_strict"]:
        report["raw_sheet"] = "完全原样保留（未叠加任何美化）"
    else:
        kinds_raw, display_raw = build_raw_sheet(ws_raw, profile, header_row, last_row,
                                                 ncols_raw, col_of)
        report["raw_sheet"] = "值/行列顺序/数字格式原样保留，已叠加表头美化、斑马纹、边框、列宽、冻结"

    # ---- 新建 7 张工作表
    for name in profile["sheet_order"]:
        if name != C.SHEET_RAW and name not in wb.sheetnames:
            wb.create_sheet(name)
    wb._sheets = [wb[n] for n in profile["sheet_order"]]

    expected = {}
    formula_cells = 0

    # 任务1
    ws = wb[C.SHEET_MERGE]
    b, headers, kinds = build_merge_sheet(ws, res, profile, raw_letters)
    finalize(ws, b, headers, kinds, profile, C.SHEET_TITLES[C.SHEET_MERGE])
    expected[C.SHEET_MERGE] = b.expected

    # 任务2
    ws = wb[C.SHEET_REASON]
    b, headers, kinds = build_reason_sheet(ws, res, profile, raw_letters)
    finalize(ws, b, headers, kinds, profile, C.SHEET_TITLES[C.SHEET_REASON])
    S.reason_conditional_formats(ws, 2, 3, len(res["reasons"]), [x["key"] for x in res["reasons"]])
    expected[C.SHEET_REASON] = b.expected
    report["reason_cf_rules"] = len(res["reasons"])

    # 统计总览
    ws = wb[C.SHEET_OVERVIEW]
    b, headers, kinds, row_kinds = build_overview_sheet(ws, res, profile, raw_letters)
    finalize(ws, b, headers, kinds, profile, C.SHEET_TITLES[C.SHEET_OVERVIEW],
             ncols=2, row_kinds=row_kinds)
    expected[C.SHEET_OVERVIEW] = b.expected

    # 任务3
    for sheet_name, dim in C.DIMENSIONS:
        ws = wb[sheet_name]
        b, headers, kinds = build_dim_sheet(ws, dim, res, profile, raw_letters)
        finalize(ws, b, headers, kinds, profile, C.SHEET_TITLES[sheet_name])
        expected[sheet_name] = b.expected

    if not profile["raw_preserve_strict"]:
        S.autofit(ws_raw, ncols_raw,
                  [ws_raw.cell(header_row, c).value or "" for c in range(1, ncols_raw + 1)],
                  [[ws_raw.cell(r, c).value for c in range(1, ncols_raw + 1)]
                   for r in range(header_row + 1, last_row + 1)],
                  ["text"] * ncols_raw, profile["money_format"])
        S.freeze(ws_raw)

    # ---- 保存 + 注入公式缓存值
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    wb.save(out)

    stats = CV.inject_cached_values(out, expected)
    report["formula_cells_written"] = stats["formula_cells"]
    report["cached_values_injected"] = stats["injected"]
    report["cached_by_sheet"] = stats["sheets"]

    errors = CV.scan_errors(out)
    report["formula_errors"] = errors[:20]
    report["formula_error_count"] = len(errors)

    if args.soffice_recalc:
        report["soffice_recalc"] = CV.recalc_with_soffice(out)

    report["sheet_order"] = [w.title for w in load_workbook(out).worksheets]
    report["elapsed_sec"] = round(time.time() - t0, 2)
    report["status"] = "success" if not errors else "success_with_formula_errors"

    text = json.dumps(report, ensure_ascii=False, indent=2, default=str)
    if args.report:
        with open(args.report, "w", encoding="utf-8") as fh:
            fh.write(text)
    print(text)
    return 0 if not errors else 2


def main():
    p = argparse.ArgumentParser(description="合同数据统计核对 Skill 主程序")
    p.add_argument("--input", required=True, help="输入《合同开票及回款核对表.xlsx》绝对路径")
    p.add_argument("--output", required=True, help="输出结果 xlsx 绝对路径")
    p.add_argument("--name", default=None, help="姓名后缀；提供时按命名规范重命名输出文件")
    p.add_argument("--profile", default=C.DEFAULT_PROFILE, choices=sorted(C.PROFILES),
                   help="spec=以题目正文为准（默认）；example=以附件2示例为准")
    p.add_argument("--raw-strict", action="store_true",
                   help="「原始数据」表完全不叠加美化，最严格的原样保留")
    p.add_argument("--soffice-recalc", action="store_true",
                   help="额外用 LibreOffice 重算做交叉验证（非必需，缺环境自动跳过）")
    p.add_argument("--report", default=None, help="执行回执 JSON 输出路径")
    args = p.parse_args()
    sys.exit(run(args))


if __name__ == "__main__":
    main()
