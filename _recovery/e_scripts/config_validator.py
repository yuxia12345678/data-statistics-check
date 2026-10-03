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
