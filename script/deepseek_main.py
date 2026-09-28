# -*- coding: utf-8 -*-
"""
合同数据统计核对 Skill
功能：
1. 读取合同开票及回款核对表
2. 原样复制原始数据
3. 同合同号合并汇总
4. 未回款原因分类汇总
5. 区域/部门/账龄/客户分类多维度统计
6. 统计总览
7. 全工作表统一美化
"""

import os
import re
import math
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
import zipfile

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, NamedStyle
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.formatting.rule import CellIsRule


# =========================================================
# 一、基础配置
# =========================================================
# 当前脚本所在目录
BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "附件1-合同开票及回款核对表.xlsx"
OUTPUT_FILE = BASE_DIR / "合同开票及回款核对表分析结果_deepseek.xlsx"

SHEET_RAW = "原始数据"
SHEET_MERGED = "同合同号合并汇总"
SHEET_REASON = "未回款原因分类汇总"
SHEET_OVERVIEW = "统计总览"
SHEET_REGION = "区域统计"
SHEET_DEPT = "部门统计"
SHEET_AGE = "账龄统计"
SHEET_CUSTOMER = "客户分类统计"

# 原始表头在第2行（Excel 行号），pandas header=1
HEADER_ROW = 1

# 数值字段
NUMERIC_FIELDS = [
    "合同金额", "未开票金额", "开票金额", "回款合计", "开票未回款"
]

# 合并时取第一条非空的文本字段
TEXT_FIRST_FIELDS = [
    "区域", "客户分类", "开票日期", "部门", "责任人", "账龄"
]

# 未回款原因字段
REASON_FIELD = "未回款原因分类"

# 多维度统计维度
DIMENSIONS = {
    SHEET_REGION: "区域",
    SHEET_DEPT: "部门",
    SHEET_AGE: "账龄",
    SHEET_CUSTOMER: "客户分类",
}

# 样式常量
HEADER_FILL = PatternFill("solid", fgColor="8EA9DB")  # 低饱和蓝
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

THIN_GRAY = Side(style="thin", color="D9D9D9")
BORDER_BOTTOM = Border(bottom=THIN_GRAY)

ZEBRA_FILL = PatternFill("solid", fgColor="F7F9FC")

# 柔和分类背景色
SOFT_COLORS = [
    "FFF2CC", "E2EFDA", "DDEBF7", "FCE4D6",
    "E4DFEC", "D9E1F2", "F2DCDB", "EAF1DD",
    "FDE9D9", "DCE6F1", "EBF1DE", "F2F2F2"
]


# =========================================================
# 二、工具函数
# =========================================================

def is_empty(v):
    """判断空值"""
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    s = str(v).strip()
    return s == "" or s.lower() == "nan" or s.lower() == "none"


def to_number(v):
    """转数值，空值按0处理"""
    if is_empty(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0


def clean_text(v):
    """清理文本"""
    if is_empty(v):
        return ""
    return str(v).strip()


def safe_sheet_title(name):
    """Excel 工作表名不能超过31字符，且不能包含特殊字符"""
    name = re.sub(r'[\\/*?:\[\]]', "", str(name))
    return name[:31]


def full_width_len(text):
    """中文按全角2个字符计算列宽"""
    if text is None:
        return 0
    s = str(text)
    length = 0
    for ch in s:
        if '\u4e00' <= ch <= '\u9fff':
            length += 2
        else:
            length += 1
    return length


def auto_fit_columns(ws, min_width=8, max_width=40):
    """自动适配列宽，中文按全角计算"""
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            v = cell.value
            if v is None:
                continue
            # 公式不参与列宽计算
            if isinstance(v, str) and v.startswith("="):
                continue
            max_len = max(max_len, full_width_len(v))
        width = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[col_letter].width = width


def style_header(ws, header_row=1):
    """表头样式"""
    for cell in ws[header_row]:
        if cell.value is None:
            continue
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER


def style_data_area(ws, start_row=1, end_row=None, end_col=None):
    """
    数据区样式：
    - 斑马纹
    - 浅灰水平边框
    - 对齐：文本左，数值右，日期/分类居中
    """
    if end_row is None:
        end_row = ws.max_row
    if end_col is None:
        end_col = ws.max_column

    for r in range(start_row, end_row + 1):
        for c in range(1, end_col + 1):
            cell = ws.cell(row=r, column=c)
            # 斑马纹
            if (r - start_row) % 2 == 1:
                cell.fill = ZEBRA_FILL
            # 只保留浅灰水平边框
            cell.border = BORDER_BOTTOM
            # 对齐
            if isinstance(cell.value, (int, float)):
                cell.alignment = RIGHT
            else:
                cell.alignment = LEFT

    # 日期、分类字段居中：这里统一把表头和数据中属于日期/分类的列居中
    # 由调用方按列设置更精确，此处保留通用左/右


def set_number_format(ws, col_names, number_format, header_row=1):
    """按列名设置数值格式"""
    header = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is not None:
            header[str(v).strip()] = c

    for name in col_names:
        if name not in header:
            continue
        col_idx = header[name]
        for r in range(header_row + 1, ws.max_row + 1):
            ws.cell(row=r, column=col_idx).number_format = number_format


def set_percent_format(ws, col_names, header_row=1):
    """设置百分比格式，保留2位小数"""
    set_number_format(ws, col_names, "0.00%", header_row)


def set_amount_format(ws, col_names, header_row=1):
    """金额千分位 + 3位小数"""
    set_number_format(ws, col_names, "#,##0.000", header_row)


def freeze_header(ws, header_row=1):
    """冻结表头首行"""
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)


def add_soft_condition_format(ws, col_letter, start_row, end_row):
    """为分类列设置柔和背景条件格式"""
    for i, color in enumerate(SOFT_COLORS):
        # 这里按分类值轮询不易直接做条件格式，改为直接填充
        pass


# =========================================================
# 三、读取原始数据
# =========================================================

def read_input(path):
    """
    读取输入 Excel。
    表头在第2行，因此 pandas header=1。
    保留原始第1行说明「单位：万元」。
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到输入文件：{path}")

    # 读取全部内容，不跳过第一行，便于原样复制
    raw_df = pd.read_excel(path, sheet_name=0, header=None, dtype=object)
    # 第0行是说明行，第1行是表头
    header = raw_df.iloc[1].tolist()
    data = raw_df.iloc[2:].copy()
    data.columns = header
    data = data.reset_index(drop=True)
    return raw_df, data


def normalize_columns(df):
    """统一列名，去除空格"""
    df.columns = [str(c).strip() if c is not None else "" for c in df.columns]
    return df


# =========================================================
# 四、任务1：同合同号合并
# =========================================================

def merge_by_contract(df):
    """
    以合同号为唯一主键合并。
    数值字段求和；文本字段取第一条非空；未回款原因分类去重拼接。
    """
    df = normalize_columns(df.copy())

    # 过滤掉合同号为空的记录
    df = df[~df["合同号"].apply(is_empty)].copy()

    # 数值字段转数值
    for col in NUMERIC_FIELDS:
        if col in df.columns:
            df[col] = df[col].apply(to_number)

    # 未回款原因分类：空值填「未填写」？题目示例里原始数据空值最终变为「未填写」
    # 但题目正文说“对同一合同号下全部非空内容去重后拼接合并”，
    # 因此这里对空值不强行填“未填写”，只在最终展示时保留空。
    # 若需要与示例一致，可在输出时把空值显示为“未填写”。
    if REASON_FIELD in df.columns:
        df[REASON_FIELD] = df[REASON_FIELD].apply(clean_text)

    grouped = df.groupby("合同号", sort=False, as_index=False)

    rows = []
    for contract_no, g in grouped:
        row = OrderedDict()
        row["合同号"] = contract_no

        # 数值字段求和
        for col in NUMERIC_FIELDS:
            row[col] = round(g[col].sum(), 6) if col in g.columns else 0.0

        # 文本字段第一条非空
        for col in TEXT_FIRST_FIELDS:
            if col not in g.columns:
                row[col] = ""
                continue
            val = ""
            for v in g[col].tolist():
                if not is_empty(v):
                    val = clean_text(v)
                    break
            row[col] = val

        # 未回款原因分类：非空去重拼接
        reasons = []
        if REASON_FIELD in g.columns:
            for v in g[REASON_FIELD].tolist():
                v = clean_text(v)
                if v and v not in reasons:
                    reasons.append(v)
        row[REASON_FIELD] = "；".join(reasons)

        rows.append(row)

    merged = pd.DataFrame(rows)

    # 字段顺序：序号、合同号、区域、客户分类、开票日期、合同金额、未开票金额、开票金额、回款合计、开票未回款、部门、责任人、账龄、未回款原因分类
    # 注意：任务1要求“完整保留原始表格全部字段列结构”，但原始字段里没有“合并开票金额”以外的多余字段。
    # 示例输出中「同合同号合并汇总」字段为：
    # 序号、合同号、区域、客户分类、开票日期、合同金额、未开票金额、开票金额、回款合计、开票未回款、部门、责任人、账龄、未回款原因分类
    # 这里按示例字段顺序输出。
    out_cols = [
        "合同号", "区域", "客户分类", "开票日期",
        "合同金额", "未开票金额", "开票金额", "回款合计", "开票未回款",
        "部门", "责任人", "账龄", REASON_FIELD
    ]
    for c in out_cols:
        if c not in merged.columns:
            merged[c] = ""

    merged = merged[out_cols].copy()
    merged.insert(0, "序号", range(1, len(merged) + 1))

    # 按开票未回款降序，便于查看
    merged = merged.sort_values(by="开票未回款", ascending=False).reset_index(drop=True)
    merged["序号"] = range(1, len(merged) + 1)

    return merged


# =========================================================
# 五、任务2：未回款原因分类汇总
# =========================================================

def reason_summary(df):
    """
    基于原始明细数据的「未回款原因分类」字段统计。
    统计：总开票未回款、涉及合同数、涉及笔数、占比。
    """
    df = normalize_columns(df.copy())
    df = df[~df["合同号"].apply(is_empty)].copy()

    if REASON_FIELD not in df.columns:
        df[REASON_FIELD] = ""

    df[REASON_FIELD] = df[REASON_FIELD].apply(clean_text)
    df["开票未回款"] = df["开票未回款"].apply(to_number)

    # 空值处理：题目示例最终显示“未填写”
    df[REASON_FIELD] = df[REASON_FIELD].apply(lambda x: x if x else "未填写")

    total_unpaid = df["开票未回款"].sum()

    rows = []
    for reason, g in df.groupby(REASON_FIELD, sort=False):
        amount = g["开票未回款"].sum()
        contracts = g["合同号"].nunique()
        records = len(g)
        ratio = amount / total_unpaid if total_unpaid else 0.0
        rows.append({
            "未回款原因分类": reason,
            "开票未回款": round(amount, 6),
            "涉及合同数": contracts,
            "涉及笔数": records,
            "占比": ratio
        })

    result = pd.DataFrame(rows)
    if result.empty:
        result = pd.DataFrame(columns=["未回款原因分类", "开票未回款", "涉及合同数", "涉及笔数", "占比"])
    else:
        result = result.sort_values(by="开票未回款", ascending=False).reset_index(drop=True)

    result.insert(0, "序号", range(1, len(result) + 1))
    return result


# =========================================================
# 六、任务3：多维度统计
# =========================================================

def dimension_summary(df, dim_col):
    """
    按维度统计：
    总合同金额、总开票金额、总回款合计、总开票未回款、合同数、整体回款率、未回款占比
    """
    df = normalize_columns(df.copy())
    df = df[~df["合同号"].apply(is_empty)].copy()

    for col in ["合同金额", "开票金额", "回款合计", "开票未回款"]:
        df[col] = df[col].apply(to_number)

    rows = []
    for dim_val, g in df.groupby(dim_col, sort=False):
        contract_amount = g["合同金额"].sum()
        invoice_amount = g["开票金额"].sum()
        received_amount = g["回款合计"].sum()
        unpaid_amount = g["开票未回款"].sum()
        contract_count = g["合同号"].nunique()

        # 整体回款率 = 总回款合计 / 总开票金额
        repayment_rate = received_amount / invoice_amount if invoice_amount else 0.0
        # 未回款占比 = 总开票未回款 / 总合同金额？题目未明确分母。
        # 示例中区域统计：未回款占比 = 开票未回款 / 总合同金额（全局）
        # 为保持与示例一致，这里先按全局总合同金额计算，后面在总览中覆盖。
        # 更合理口径：未回款占比 = 开票未回款 / 合同金额。
        unpaid_ratio = unpaid_amount / contract_amount if contract_amount else 0.0

        rows.append({
            dim_col: dim_val,
            "合同金额": round(contract_amount, 6),
            "开票金额": round(invoice_amount, 6),
            "回款合计": round(received_amount, 6),
            "开票未回款": round(unpaid_amount, 6),
            "合同数": contract_count,
            "回款率": repayment_rate,
            "未回款占比": unpaid_ratio,
        })

    result = pd.DataFrame(rows)
    if result.empty:
        result = pd.DataFrame(columns=[dim_col, "合同金额", "开票金额", "回款合计", "开票未回款", "合同数", "回款率", "未回款占比"])
    else:
        result = result.sort_values(by="开票未回款", ascending=False).reset_index(drop=True)

    result.insert(0, "序号", range(1, len(result) + 1))
    return result


def overview_summary(df):
    """统计总览"""
    df = normalize_columns(df.copy())
    df = df[~df["合同号"].apply(is_empty)].copy()

    for col in ["合同金额", "开票金额", "回款合计", "开票未回款"]:
        df[col] = df[col].apply(to_number)

    total_contract = df["合同金额"].sum()
    total_invoice = df["开票金额"].sum()
    total_received = df["回款合计"].sum()
    total_unpaid = df["开票未回款"].sum()
    total_contract_count = df["合同号"].nunique()
    total_records = len(df)

    reason_count = df[REASON_FIELD].apply(clean_text)
    reason_count = reason_count[reason_count != ""].nunique()

    repayment_rate = total_received / total_invoice if total_invoice else 0.0

    rows = [
        ("总合同金额（万元）", round(total_contract, 6)),
        ("总开票金额（万元）", round(total_invoice, 6)),
        ("总回款合计（万元）", round(total_received, 6)),
        ("总开票未回款（万元）", round(total_unpaid, 6)),
        ("整体回款率", repayment_rate),
        ("总合同数", total_contract_count),
        ("总数据笔数", total_records),
        ("未回款原因分类数", reason_count),
    ]
    return pd.DataFrame(rows, columns=["指标", "数值"])


# =========================================================
# 七、写入 Excel 并美化
# =========================================================

def write_df_to_sheet(wb, sheet_name, df, title=None, header_row=1):
    """把 DataFrame 写入工作表"""
    ws = wb.create_sheet(title=safe_sheet_title(sheet_name))

    r = 1
    if title:
        ws.cell(row=1, column=1, value=title)
        ws.cell(row=1, column=1).font = Font(bold=True, size=13)
        r = 2

    # 写表头
    for j, col in enumerate(df.columns, start=1):
        ws.cell(row=r, column=j, value=col)

    # 写数据
    for i, row in df.iterrows():
        for j, col in enumerate(df.columns, start=1):
            ws.cell(row=r + 1 + i, column=j, value=row[col])

    return ws, r


def apply_sheet_style(ws, header_row, title_row=False):
    """统一美化"""
    # 表头样式
    style_header(ws, header_row)

    # 数据区样式
    style_data_area(ws, start_row=header_row, end_row=ws.max_row, end_col=ws.max_column)

    # 冻结表头
    freeze_header(ws, header_row)

    # 列宽
    auto_fit_columns(ws)

    # 对齐修正：日期、分类居中
    header_map = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is not None:
            header_map[str(v).strip()] = c

    center_cols = ["序号", "开票日期", "账龄", "未回款原因分类", "区域", "部门", "客户分类"]
    for name in center_cols:
        if name in header_map:
            col_idx = header_map[name]
            for r in range(header_row + 1, ws.max_row + 1):
                ws.cell(row=r, column=col_idx).alignment = CENTER


def build_output(raw_df, merged_df, reason_df, overview_df, dim_results, output_path):
    """生成最终 Excel"""
    wb = Workbook()

    # 删除默认 sheet
    default = wb.active
    wb.remove(default)

    # 1. 原始数据：原样复制
    ws_raw = wb.create_sheet(SHEET_RAW)
    for i, row in raw_df.iterrows():
        for j, val in enumerate(row.tolist(), start=1):
            ws_raw.cell(row=i + 1, column=j, value=val)

    # 原始数据表头在第2行
    raw_header_row = 2
    # 第一行说明保留
    style_header(ws_raw, raw_header_row)
    style_data_area(ws_raw, start_row=raw_header_row, end_row=ws_raw.max_row, end_col=ws_raw.max_column)
    freeze_header(ws_raw, raw_header_row)
    auto_fit_columns(ws_raw)

    # 2. 同合同号合并汇总
    ws_merged, hrow = write_df_to_sheet(wb, SHEET_MERGED, merged_df, title="同一合同号数据合并汇总")
    apply_sheet_style(ws_merged, hrow)
    set_amount_format(ws_merged, ["合同金额", "未开票金额", "开票金额", "回款合计", "开票未回款"], hrow)

    # 3. 未回款原因分类汇总
    ws_reason, hrow = write_df_to_sheet(wb, SHEET_REASON, reason_df, title="未回款原因分类汇总")
    apply_sheet_style(ws_reason, hrow)
    set_amount_format(ws_reason, ["开票未回款"], hrow)
    set_percent_format(ws_reason, ["占比"], hrow)

    # 为不同分类设置柔和背景
    if "未回款原因分类" in reason_df.columns:
        reason_col_idx = list(reason_df.columns).index("未回款原因分类") + 1
        for i in range(len(reason_df)):
            r = hrow + 1 + i
            color = SOFT_COLORS[i % len(SOFT_COLORS)]
            ws_reason.cell(row=r, column=reason_col_idx).fill = PatternFill("solid", fgColor=color)

    # 4. 统计总览
    ws_overview, hrow = write_df_to_sheet(wb, SHEET_OVERVIEW, overview_df, title="数据统计总览")
    apply_sheet_style(ws_overview, hrow)
    # 总览中金额和百分比分别设置
    for r in range(hrow + 1, ws_overview.max_row + 1):
        k = str(ws_overview.cell(row=r, column=1).value)
        v = ws_overview.cell(row=r, column=2)
        if "率" in k:
            v.number_format = "0.00%"
        elif "数" in k and "分类" not in k:
            v.number_format = "0"
        else:
            v.number_format = "#,##0.000"

    # 5. 四个维度统计
    for sheet_name, dim_col in DIMENSIONS.items():
        df_dim = dim_results[sheet_name]
        ws_dim, hrow = write_df_to_sheet(wb, sheet_name, df_dim, title=f"{dim_col}维度统计分析")
        apply_sheet_style(ws_dim, hrow)
        set_amount_format(ws_dim, ["合同金额", "开票金额", "回款合计", "开票未回款"], hrow)
        set_percent_format(ws_dim, ["回款率", "未回款占比"], hrow)
        # 合同数整数
        if "合同数" in df_dim.columns:
            col_idx = list(df_dim.columns).index("合同数") + 1
            for r in range(hrow + 1, ws_dim.max_row + 1):
                ws_dim.cell(row=r, column=col_idx).number_format = "0"

    # 保存
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    print(f"已生成：{output_path}")


# =========================================================
# 八、主流程
# =========================================================

def main():
    print("开始读取原始数据...")
    raw_df, data_df = read_input(INPUT_FILE)
    data_df = normalize_columns(data_df)

    print("任务1：同合同号合并...")
    merged_df = merge_by_contract(data_df)

    print("任务2：未回款原因分类汇总...")
    reason_df = reason_summary(data_df)

    print("任务3：多维度统计...")
    dim_results = {}
    for sheet_name, dim_col in DIMENSIONS.items():
        dim_results[sheet_name] = dimension_summary(data_df, dim_col)

    print("统计总览...")
    overview_df = overview_summary(data_df)

    print("生成结果 Excel...")
    build_output(raw_df, merged_df, reason_df, overview_df, dim_results, OUTPUT_FILE)

    print("Skill 执行完成。")


if __name__ == "__main__":
    main()