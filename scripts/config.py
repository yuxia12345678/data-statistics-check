# -*- coding: utf-8 -*-
"""
常量与 profile 定义（单一事实源）。

本文件是 `verify.py` 的依赖：交付前自检用它提供的字段类型表（FIELD_KIND）、工作表名与
顺序（SHEET_TITLES / DIMENSIONS / sheet_order）、数字格式与样式常量，做 44 条逐格断言
（R1、R5–R16）。

⚠️ 产物生成侧的规则与样式**不在本文件**，而在 `config/contract_repayment.json`
（由 `run_skill.py → stat_engine.py` 读取）。两侧必须保持一致，否则自检会报 R8–R16 不合规。

规范来源：
  - 业务口径与公式  见 `references/口径与公式规范.md`
  - 样式可执行参数  见 `references/样式规范.md`
  - spec/example 取舍 见 `references/与示例文件的差异.md`

两个 profile：
  - `spec`    （默认）以《题目描述》正文为评分规范；
  - `example` 以《附件2-示例》形态为准（金额 2 位小数、四边框、日期右对齐、
              分类左对齐、「原始数据」第 1 行覆盖为大标题）。
"""

import copy


# ================================================================ 工作表

SHEET_RAW = "原始数据"
SHEET_MERGE = "同合同号合并汇总"
SHEET_REASON = "未回款原因分类汇总"
SHEET_OVERVIEW = "统计总览"

# (工作表名, 维度字段名)，顺序即工作表输出顺序
DIMENSIONS = [
    ("区域统计", "区域"),
    ("部门统计", "部门"),
    ("账龄统计", "账龄"),
    ("客户分类统计", "客户分类"),
]

# 各表第 1 行跨列合并大标题（取自《附件2-示例》）
SHEET_TITLES = {
    SHEET_RAW: "原始数据明细",
    SHEET_MERGE: "同一合同号数据合并汇总",
    SHEET_REASON: "未回款原因分类汇总",
    SHEET_OVERVIEW: "数据统计总览",
    "区域统计": "区域维度统计分析",
    "部门统计": "部门维度统计分析",
    "账龄统计": "账龄维度统计分析",
    "客户分类统计": "客户分类维度统计分析",
}

DEFAULT_PROFILE = "spec"


# ================================================================ 字段口径

# 题目《完整字段清单》14 字段（「原始数据」另含源文件多出的 G 列「合并开票金额」）
TASK_FIELDS = [
    "序号", "区域", "客户分类", "合同号", "开票日期",
    "合同金额", "未开票金额", "开票金额", "回款合计", "开票未回款",
    "部门", "责任人", "账龄", "未回款原因分类",
]

# 任务1：同合同号合并时，数值字段求和
MERGE_SUM_FIELDS = ["合同金额", "未开票金额", "开票金额", "回款合计", "开票未回款"]
# 任务1：普通文本字段取分组第一条非空有效值
MERGE_FIRST_FIELDS = ["区域", "客户分类", "开票日期", "部门", "责任人", "账龄"]
# 任务1：备注字段组内非空去重后拼接
MERGE_JOIN_FIELDS = ["未回款原因分类"]
JOIN_SEP = "、"

# 任务2：未回款原因分类汇总表头（6 列，含序号）
REASON_HEADERS = ["序号", "未回款原因分类", "开票未回款", "涉及合同数", "涉及数据笔数", "占比"]

# 统计总览 8 项全局指标标签（顺序与数值一一对应）
OVERVIEW_LABELS = [
    "总合同金额（万元）",
    "总开票金额（万元）",
    "总回款合计（万元）",
    "总开票未回款（万元）",
    "整体回款率",
    "总合同数",
    "总数据笔数",
    "未回款原因分类数",
]

# 「原始数据」表数据起始行（表头行由程序动态探测，源文件表头在第 2 行）
SRC_DATA_FIRST_ROW = 3

# 字段类型表：决定数字格式与对齐（数值右、文本左、日期与分类居中）
# 同一字段的不同写法（带「总」/ 不带「总」）都要登记，供 verify.R8/R9 反查。
FIELD_KIND = {
    # 通用
    "序号": "int",
    "指标": "text",
    "数值": "num",
    # 明细 / 合并表
    "区域": "cat",
    "客户分类": "cat",
    "合同号": "text",
    "开票日期": "date",
    "合同金额": "money",
    "合并开票金额": "money",
    "未开票金额": "money",
    "开票金额": "money",
    "回款合计": "money",
    "开票未回款": "money",
    "部门": "cat",
    "责任人": "text",
    "账龄": "cat",
    "未回款原因分类": "cat",
    # 任务3 维度表（spec 表头带「总」字）
    "总合同金额": "money",
    "总开票金额": "money",
    "总回款合计": "money",
    "总开票未回款": "money",
    "合同数": "int",
    "整体回款率": "pct",
    "未回款占比": "pct",
    # 任务3 维度表（example 表头不带「总」字）
    "回款率": "pct",
    # 任务2 原因表
    "涉及合同数": "int",
    "涉及数据笔数": "int",
    "占比": "pct",
}


# ================================================================ 样式常量
# 全部取自《样式规范》「可执行参数」，verify.R8–R16 逐格断言这些值。

TITLE_FONT_NAME = "微软雅黑"
TITLE_FONT_SIZE = 14
TITLE_FONT_COLOR = "FFFFFFFF"
TITLE_FILL = "FF1F4E78"
TITLE_ROW_HEIGHT = 37.35

HEADER_FONT_NAME = "微软雅黑"
HEADER_FONT_SIZE = 11
HEADER_FONT_COLOR = "FFFFFFFF"
HEADER_FILL = "FF0070C0"
HEADER_ROW_HEIGHT = 32.45

DATA_FONT_NAME = "微软雅黑"
DATA_FONT_SIZE = 11
DATA_FONT_COLOR = "FF000000"

ZEBRA_FILL = "FFEBF1F8"
BORDER_COLOR = "FFBFBFBF"

MONEY_FORMAT_SPEC = "#,##0.000"     # 题目正文：千分位 + 3 位小数
MONEY_FORMAT_EXAMPLE = "#,##0.00"   # 附件2-示例：2 位小数
PCT_FORMAT = "0.00%"
INT_FORMAT = "0"
DATE_FORMAT_SRC = "@"               # 开票日期维持源文本 yyyymmdd

COL_WIDTH_MIN = 8
COL_WIDTH_MAX = 40
COL_WIDTH_PADDING = 3

FREEZE_PANES = "A3"

# 任务2 差异化柔和背景（循环使用）
REASON_SOFT_COLORS = [
    "FFFDE9D9", "FFDCE6F1", "FFEBF1DE", "FFE4DFEC",
    "FFF2DCDB", "FFDAEEF3", "FFFFF2CC", "FFEAF1DD",
]

# 任务1 合并表列序 = 题目 14 字段清单（源文件列序，剔除多出的「合并开票金额」）
_MERGE_COL_ORDER_SPEC = [
    "序号", "区域", "客户分类", "合同号", "开票日期",
    "合同金额", "未开票金额", "开票金额", "回款合计", "开票未回款",
    "部门", "责任人", "账龄", "未回款原因分类",
]

# 任务3 维度表表头（`{dim}` 占位符替换为维度名）
_DIM_HEADERS_SPEC = [
    "序号", "{dim}", "总合同金额", "总开票金额", "总回款合计",
    "总开票未回款", "合同数", "整体回款率", "未回款占比",
]
_DIM_HEADERS_EXAMPLE = [
    "序号", "{dim}", "合同金额", "开票金额", "回款合计",
    "开票未回款", "合同数", "回款率", "未回款占比",
]

# 统计总览排最后（2026-10-03 由第 4 位改为最后，与示例一致）
_SHEET_ORDER_SPEC = [
    SHEET_RAW, SHEET_MERGE, SHEET_REASON,
    "区域统计", "部门统计", "账龄统计", "客户分类统计", SHEET_OVERVIEW,
]
_SHEET_ORDER_EXAMPLE = [
    SHEET_RAW, SHEET_MERGE, SHEET_REASON,
    "区域统计", "部门统计", "账龄统计", "客户分类统计", SHEET_OVERVIEW,
]

# 对齐：文本左、数值右、日期与分类居中（题目正文）
_ALIGN_SPEC = {
    "text": "left", "num": "right", "money": "right",
    "int": "right", "pct": "right", "date": "center", "cat": "center",
}
# 附件2-示例：开票日期右对齐、分类字段左对齐
_ALIGN_EXAMPLE = dict(_ALIGN_SPEC, date="right", cat="left")

PROFILES = {
    "spec": {
        "name": "spec",
        "money_format": MONEY_FORMAT_SPEC,
        "border_mode": "horizontal",       # 仅浅灰水平边框
        "align_by_kind": dict(_ALIGN_SPEC),
        "merge_col_order": list(_MERGE_COL_ORDER_SPEC),
        "dim_headers": list(_DIM_HEADERS_SPEC),
        "sheet_order": list(_SHEET_ORDER_SPEC),
        "raw_override_title_row": False,   # 保留源第 1 行「单位：万元」
        "raw_preserve_strict": False,      # True 时「原始数据」不叠加任何美化
        "fill_blank_reason": False,        # 空分类不视为一个分类
    },
    "example": {
        "name": "example",
        "money_format": MONEY_FORMAT_EXAMPLE,
        "border_mode": "all",              # 四边 thin 全边框
        "align_by_kind": dict(_ALIGN_EXAMPLE),
        "merge_col_order": list(_MERGE_COL_ORDER_SPEC),
        "dim_headers": list(_DIM_HEADERS_EXAMPLE),
        "sheet_order": list(_SHEET_ORDER_EXAMPLE),
        "raw_override_title_row": True,    # 覆盖第 1 行为合并大标题
        "raw_preserve_strict": False,
        "fill_blank_reason": False,
    },
}


def get_profile(name=None):
    """返回 profile 的深拷贝，调用方可安全修改（如 --raw-strict）。"""
    key = name or DEFAULT_PROFILE
    if key not in PROFILES:
        raise ValueError("未知 profile：%s（可选 %s）" % (key, "、".join(sorted(PROFILES))))
    return copy.deepcopy(PROFILES[key])
