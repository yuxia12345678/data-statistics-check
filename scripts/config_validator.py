# -*- coding: utf-8 -*-
"""
配置校验模块 config_validator
职责：
1、加载JSON业务配置后，运行业务逻辑前做schema校验；
2、检查必填节点、字段，提前抛出异常，避免中途运行报错；
3、严格校验8个工作表名称与顺序，与配置声明的期望值不匹配直接抛出ValueError；
4、只校验配置结构，不校验业务数据。

可配置性说明（对应《竞赛须知》"Skill规则可配置要求"）
----------------------------------------------------
本模块**不写死任何业务表名**：期望的表名与顺序取自
`profiles.<default_profile>.sheet_order`，与实际输出表名（`output_sheet_names` +
`task2_special_agg.sheet_name` + `task3_multi_dim.overview_sheet_name` + `dim_list`）
双向核对。因此换一套业务配置 = 换一套表名与规则，代码零改动。
`DEFAULT_SPEC_SHEET_ORDER` 仅是配置未声明时的兜底（即本题《数据统计核对》规定的 8 张表）。
"""
from typing import Dict, Any

# 本题（数据统计核对）规定的 8 张工作表名称与顺序 —— 仅作**兜底默认值**。
# 配置里 `profiles.<default_profile>.sheet_order` 一旦声明，期望值以配置为准。
DEFAULT_SPEC_SHEET_ORDER = [
    "原始数据",
    "同合同号合并汇总",
    "未回款原因分类汇总",
    "统计总览",
    "区域统计",
    "部门统计",
    "账龄统计",
    "客户分类统计",
]

# 期望的工作表数量（题目硬性要求：必须包含 8 张独立工作表）
EXPECT_SHEET_COUNT = 8


def expected_sheet_order(cfg: Dict[str, Any]):
    """返回期望的 8 张工作表**名称与顺序**。

    唯一事实源：`profiles.<default_profile>.sheet_order`；
    未声明时回落到 `DEFAULT_SPEC_SHEET_ORDER`（本题口径）。
    """
    profiles = cfg.get("profiles") or {}
    prof = profiles.get(cfg.get("default_profile")) if cfg.get("default_profile") else None
    if not isinstance(prof, dict):
        prof = next((p for p in profiles.values() if isinstance(p, dict)), None) or {}
    order = prof.get("sheet_order")
    return list(order) if order else list(DEFAULT_SPEC_SHEET_ORDER)


def validate_business_config(cfg: Dict[str, Any]) -> None:
    """
    业务配置整体校验入口
    :param cfg: json.load读取后的配置字典
    :return: 无返回，校验失败抛ValueError
    """
    # 顶层必填节点
    must_top_keys = [
        "input", "style_setting", "output_filename", "output_sheet_names",
        "task1_group_merge", "task2_special_agg", "task3_multi_dim"
    ]
    for k in must_top_keys:
        if k not in cfg:
            raise ValueError(f"配置缺失顶层节点：{k}")

    input_cfg = cfg["input"]
    for k in ["sheet_name", "header_row"]:
        if k not in input_cfg:
            raise ValueError(f"input节点缺失字段：{k}")

    # 任务1校验
    t1 = cfg["task1_group_merge"]
    if "primary_key" not in t1 or "agg_strategy_list" not in t1:
        raise ValueError("task1_group_merge 需要配置 primary_key、agg_strategy_list")

    # 任务2校验
    t2 = cfg["task2_special_agg"]
    for k in ["group_field", "amount_field", "contract_key", "sheet_name",
              "metrics", "target_sum_field", "ratio_field_name", "sort_by_field"]:
        if k not in t2:
            raise ValueError(f"task2_special_agg缺失配置项：{k}")

    # 任务3校验
    t3 = cfg["task3_multi_dim"]
    for k in ["dim_list", "overview_sheet_name", "dim_metrics",
              "derived_ratio_list", "sort_by_field", "overview_metrics"]:
        if k not in t3:
            raise ValueError(f"task3_multi_dim缺失配置项：{k}")

    # 工作表名校验：期望值来自配置（profiles.<default_profile>.sheet_order），
    # 而非写死在代码里 —— 换案例/换规则时随配置一起替换，代码零改动。
    out_sheet = cfg["output_sheet_names"]
    for k in ["raw_copy", "task1_result"]:
        if k not in out_sheet:
            raise ValueError(f"output_sheet_names缺失配置项：{k}")

    sheet_collect = [
        out_sheet["raw_copy"],
        out_sheet["task1_result"],
        t2["sheet_name"],
        t3["overview_sheet_name"]
    ]
    for dim_item in t3["dim_list"]:
        sheet_collect.append(dim_item["sheet_name"])

    expect_sheet_order = expected_sheet_order(cfg)
    if len(expect_sheet_order) != EXPECT_SHEET_COUNT:
        raise ValueError(
            f"期望工作表数量错误！必须为 {EXPECT_SHEET_COUNT} 张，"
            f"实际声明 {len(expect_sheet_order)} 张：{expect_sheet_order}"
        )
    if len(sheet_collect) != EXPECT_SHEET_COUNT or len(set(sheet_collect)) != EXPECT_SHEET_COUNT:
        raise ValueError(
            f"工作表数量或重名错误！必须为 {EXPECT_SHEET_COUNT} 张互不重名的表，实际:{sheet_collect}"
        )
    if sheet_collect != expect_sheet_order:
        missing = [s for s in expect_sheet_order if s not in sheet_collect]
        extra = [s for s in sheet_collect if s not in expect_sheet_order]
        raise ValueError(
            "工作表名称/顺序错误！名字、文字、大小写、先后顺序必须完全一致\n"
            f"  期望:{expect_sheet_order}\n"
            f"  实际:{sheet_collect}\n"
            f"  缺失:{missing} 多余:{extra}"
        )


# ================================================================ 字段类型清单校验

# 驱动对齐与数字格式的字段类型清单键（与 excel_styler.apply_sheet_format 一一对应）
FIELD_KIND_KEYS = [
    "num_fields", "int_fields", "pct_fields", "date_fields", "cat_fields", "text_fields"
]


def validate_sheet_field_kinds(sheet_name: str,
                               headers,
                               kind_lists: Dict[str, Any],
                               kind_keys=None) -> None:
    """
    校验一张工作表的**表头是否全部登记了字段类型**。

    对齐规则（文本左 / 数值右 / 日期与分类居中）与数字格式都由这些清单驱动：
    num_fields / int_fields / pct_fields / date_fields / cat_fields / text_fields。

    若有表头既不在上述任何清单中，说明该列类型未登记——它会被静默地兜底成文本
    （左对齐、不设数字格式），很可能不是预期结果。因此这里**直接抛 ValueError**，
    强制配置方显式登记每一列。

    :param sheet_name: 工作表名（仅用于错误信息）
    :param headers:    表头名称序列（None / 空字符串会被忽略）
    :param kind_lists: 形如 {"num_fields": [...], "text_fields": [...], ...}
    :param kind_keys:  要检查的清单键；默认 FIELD_KIND_KEYS
    :return: 无返回；存在未登记表头时抛 ValueError
    """
    keys = list(kind_keys) if kind_keys else list(FIELD_KIND_KEYS)
    classified = set()
    for k in keys:
        classified.update(kind_lists.get(k) or [])
    unknown = [h for h in headers if h not in (None, "") and h not in classified]
    if unknown:
        raise ValueError(
            "工作表「%s」有表头未登记字段类型：%s\n"
            "请在配置的 %s 中登记这些列（纯文本列请放入 text_fields）。"
            % (sheet_name, unknown, " / ".join(keys))
        )
