# -*- coding: utf-8 -*-
"""
聚合策略注册表模块
职责：
1、注册所有支持的数据聚合处理函数；
2、提供基础聚合：sum/max/min/avg/count/nunique；
3、提供特殊文本聚合：first_not_null、last_not_null、distinct_join(带分隔符文本去重拼接)；
4、本模块通用能力层，禁止写业务字段；新增聚合策略仅在此处注册，配置即可调用。
说明：本模块用于内存预处理获取维度去重值，真正Excel输出使用formula_builder生成工作表公式。
"""
import pandas as pd


def first_not_null(series: pd.Series):
    """分组取第一条非空有效值；全空返回空字符串"""
    s_valid = series.dropna()
    if len(s_valid) > 0:
        return s_valid.iloc[0]
    return ""


def last_not_null(series: pd.Series):
    """分组取最后一条非空有效值；全空返回空字符串"""
    s_valid = series.dropna()
    if len(s_valid) > 0:
        return s_valid.iloc[-1]
    return ""


def distinct_join(series: pd.Series, separator: str):
    """
    分组文本去重拼接
    过滤nan、空字符串、全空格；无有效内容返回空字符串
    """
    s_trim = series.dropna().astype(str).str.strip()
    s_valid = s_trim[s_trim != ""].unique().tolist()
    if not s_valid:
        return ""
    return separator.join(s_valid)


# 全局聚合策略注册表，配置文件通过key引用
AGG_STRATEGY_REGISTRY = {
    "sum": lambda ser: ser.sum(),
    "max": lambda ser: ser.max(),
    "min": lambda ser: ser.min(),
    "avg": lambda ser: ser.mean(),
    "first_non_null": first_not_null,
    "last_non_null": last_not_null,
    "count": lambda ser: ser.count(),
    "nunique": lambda ser: ser.nunique()
}
