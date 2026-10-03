# -*- coding: utf-8 -*-
"""
配置校验模块 config_validator
职责：
  1、加载业务json配置之后，正式运行统计逻辑之前执行schema校验；
  2、检查必填节点、必填字段是否存在；提前抛出明确的异常信息；
  3、避免引擎运行中途因为配置缺失报错；只做结构校验，不校验业务数据内容。
"""
from typing import Dict, Any


def validate_business_config(cfg: Dict[str, Any]) -> None:
    """
    业务配置整体校验入口函数
    :param cfg: 已经json.load之后的配置字典对象
    :return: None；校验失败直接抛出ValueError
    """
    # --------输入数据源配置校验--------
    if "input" not in cfg:
        raise ValueError("配置缺失input节点，输入数据源配置不能为空")
    input_cfg = cfg["input"]
    for k in ["sheet_name", "header_row"]:
        if k not in input_cfg:
            raise ValueError(f"input配置缺失字段：{k}")

    # --------任务1：主键行合并配置校验--------
    if "task1_group_merge" not in cfg:
        raise ValueError("配置缺失task1_group_merge节点（主键行合并任务）")
    t1 = cfg["task1_group_merge"]
    if "primary_key" not in t1 or "agg_strategy_list" not in t1:
        raise ValueError("task1_group_merge必须配置primary_key与agg_strategy_list")

    # --------任务2：专项分类汇总配置校验--------
    if "task2_special_agg" not in cfg:
        raise ValueError("配置缺失task2_special_agg节点（专项分类汇总任务）")
    t2 = cfg["task2_special_agg"]
    for k in ["group_field", "amount_field", "contract_key", "sheet_name"]:
        if k not in t2:
            raise ValueError(f"task2_special_agg缺失配置项：{k}")

    # --------任务3：多维度统计与全局总览校验--------
    if "task3_multi_dim" not in cfg:
        raise ValueError("配置缺失task3_multi_dim节点（多维度分析任务）")
    t3 = cfg["task3_multi_dim"]
    if "dim_list" not in t3 or "overview_sheet_name" not in t3:
        raise ValueError("task3_multi_dim必须配置dim_list以及overview_sheet_name")

    # --------工作表输出名称配置校验--------
    if "output_sheet_names" not in cfg:
        raise ValueError("配置缺失output_sheet_names节点工作表名称配置")
    out_sheet = cfg["output_sheet_names"]
    if "raw_copy" not in out_sheet or "task1_result" not in out_sheet:
        raise ValueError("output_sheet_names必须配置raw_copy、task1_result")

    # --------全局样式与输出文件名校验--------
    if "style_setting" not in cfg:
        raise ValueError("配置缺失style_setting统一可视化样式配置节点")
    if "output_filename" not in cfg:
        raise ValueError("顶层配置缺失output_filename输出文件名")
