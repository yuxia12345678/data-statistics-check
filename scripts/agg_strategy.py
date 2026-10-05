# -*- coding: utf-8 -*-
"""
聚合策略注册表 agg_strategy

职责
----
1、集中注册所有支持的数据聚合函数，配置里的 `agg_strategy` / `agg_operator` 通过 key 引用，
   新增聚合策略只需在此注册，业务代码与配置结构零改动；
2、基础聚合：sum / max / min / avg / count / nunique；
3、文本聚合：first_non_null / last_non_null / distinct_join（去重拼接）。

空白感知（重要）
----------------
源表空单元格经 pandas 读入是 `NaN`（不是 None、也不是空串），且部分字段存在
纯空白字符串。因此文本类聚合**必须**同时识别 None / NaN / NaT / 纯空白串，
否则 `str(nan)` 会以字面量 "nan" 混进结果（例如合并表「未回款原因分类」的拼接）。
本模块的 `is_blank` 是全工程唯一的空白判定实现，stat_engine 直接复用——
公式缓存值（stat_engine）与独立复核（verify.py）依赖同一语义，避免两套实现漂移。

边界说明
--------
本模块只做**内存聚合**，结果用于：公式缓存值注入、分组排序键、期望值计算；
真正写入 Excel 的是 formula_builder 生成的动态公式，两条路径互不复用。
"""
import pandas as pd


def is_blank(v) -> bool:
    """
    空白判定：None、NaN / NaT / pd.NA、纯空白字符串均视为“缺失”。

    注意：`pd.isna` 作用于数组 / 列表时返回数组而非布尔值，这里通过
    TypeError / ValueError 兜底为“不缺失”（标量场景不会走到该分支）。
    """
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip()
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def text_of(v) -> str:
    """规范化分组键 / 排序键：缺失 → 空串；其余保留源文本原样、仅去首尾空白。"""
    if is_blank(v):
        return ""
    return (v if isinstance(v, str) else str(v)).strip()


def first_non_null(series: pd.Series):
    """分组内第一条非空有效值（原始值原样返回，不做类型转换）；全空返回空字符串。"""
    for v in series.tolist():
        if not is_blank(v):
            return v
    return ""


def last_non_null(series: pd.Series):
    """分组内最后一条非空有效值；全空返回空字符串。"""
    for v in reversed(series.tolist()):
        if not is_blank(v):
            return v
    return ""


def distinct_join(series: pd.Series, separator: str = "；") -> str:
    """
    分组内非空内容去重拼接，保持**首次出现顺序**；全空返回空字符串。

    去重按“去首尾空白后的文本”比较，但拼接时同样使用规范化文本，
    与源表「同组非空内容去重后拼接」的业务口径一致。
    """
    seen = []
    for v in series.tolist():
        if is_blank(v):
            continue
        t = text_of(v)
        if t not in seen:
            seen.append(t)
    return separator.join(seen)


# 全局聚合策略注册表：配置文件通过 key 引用，新增策略仅需在此登记
AGG_STRATEGY_REGISTRY = {
    "sum": lambda ser: ser.sum(),
    "max": lambda ser: ser.max(),
    "min": lambda ser: ser.min(),
    "avg": lambda ser: ser.mean(),
    "first_non_null": first_non_null,
    "last_non_null": last_non_null,
    "count": lambda ser: ser.count(),
    "nunique": lambda ser: ser.nunique(),
}
