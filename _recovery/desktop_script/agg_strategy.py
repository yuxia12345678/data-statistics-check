# -*- coding: utf-8 -*-
"""
聚合策略注册表模块
职责：
  1、注册所有支持的数据聚合处理函数；
  2、提供基础聚合：sum/max/min/avg/count/nunique；
  3、提供特殊文本聚合：first_not_null、last_not_null、distinct_join(带分隔符文本去重拼接)；
  4、本模块属于通用能力层，禁止写任何业务字段与业务逻辑；新增聚合能力在此处注册即可。
"""
import pandas as pd


def first_not_null(series: pd.Series):
    """
    分组内取第一条非空有效值
    :param series: pandas分组后的Series对象
    :return: 返回第一个有效值；全部为空时返回空字符串，规避np.nan造成的数据类型错乱
    """
    s_valid = series.dropna()
    if len(s_valid) > 0:
        return s_valid.iloc[0]
    return ""


def last_not_null(series: pd.Series):
    """
    分组内取最后一条非空有效值
    :param series: pandas分组后的Series对象
    :return: 返回最后一个有效值；全部为空返回空字符串
    """
    s_valid = series.dropna()
    if len(s_valid) > 0:
        return s_valid.iloc[-1]
    return ""


def distinct_join(series: pd.Series, separator: str):
    """
    分组内文本去重之后按指定分隔符拼接字符串
    :param series: pandas分组Series
    :param separator: 拼接分隔符，例如 "；"
    :return: 拼接之后的文本；无有效数据返回空字符串
    """
    # 先剔除空值，再取唯一值
    s_valid = series.dropna().unique().tolist()
    if not s_valid:
        return ""
    return separator.join([str(it) for it in s_valid])


# 全局聚合策略注册表
# key:配置文件中使用的策略标识；value:对应的处理函数
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
