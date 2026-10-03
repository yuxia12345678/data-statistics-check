# -*- coding: utf-8 -*-
"""
配置校验模块 config_validator
职责：
1、加载JSON业务配置后，运行业务逻辑前做schema校验；
2、检查必填节点、字段，提前抛出异常，避免中途运行报错；
3、严格校验8个工作表名称，名字不匹配直接抛出ValueError；
4、只校验配置结构，不校验业务数据。
"""
from typing import Dict, Any


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

    # 工作表名校验：严格匹配题目8个sheet名称
    expect_sheet_set = {
        "原始数据",
        "同合同号合并汇总",
        "未回款原因分类汇总",
        "统计总览",
        "区域统计",
        "部门统计",
        "账龄统计",
        "客户分类统计"
    }
    out_sheet = cfg["output_sheet_names"]
    sheet_collect = [
        out_sheet["raw_copy"],
        out_sheet["task1_result"],
        t2["sheet_name"],
        t3["overview_sheet_name"]
    ]
    for dim_item in t3["dim_list"]:
        sheet_collect.append(dim_item["sheet_name"])
    actual_sheet_set = set(sheet_collect)
    if actual_sheet_set != expect_sheet_set:
        raise ValueError(
            f"工作表名称错误！期望:{expect_sheet_set}\n实际:{actual_sheet_set}，名字、文字、大小写必须完全一致"
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
