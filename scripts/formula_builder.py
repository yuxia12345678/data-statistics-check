# -*- coding: utf-8 -*-
"""
Excel 动态公式生成器 formula_builder
核心：全部输出 Excel 原生公式，不硬编码业务数值，数据源固定为【原始数据】工作表。

与旧版的区别（本次修正）：
1. **区间限定**：不再使用整列引用（`原始数据!$F:$F`），改为由调用方传入数据行区间
   （`first_row`/`last_row`），避免把大标题行、表头行、"单位"行纳入计算区间；
2. **除法不用 IFERROR**：统一写成 `IF(分母=0,"",分子/分母)`，语义等价且便于程序化核验；
3. **分组去重计数**：使用 `SUMPRODUCT((维度=值)/COUNTIFS(合同号,合同号&"",维度,维度))`，
   口径为"组内不同合同号的个数"（不是行数，也不是合并表行数）；
4. **首条非空**：用 `INDEX+MATCH(1,INDEX((条件)*(值<>""),0),0)` 取第一条**非空**值，
   而不是简单 MATCH 取第一条（后者在该行该字段为空时会取到空值）；
5. 新增序号、全表求和/计数、全表去重计数等构造器。

说明：本模块只拼公式字符串，不做实际内存计算；实际期望值由 stat_engine 用 pandas
独立算出，仅用于给公式注入缓存值，两条路径互不复用。
"""

RAW = "原始数据"


def _rng(col: str, first_row: int, last_row, sheet: str = RAW) -> str:
    """构造逐列区间；last_row 为 None 时退化为整列引用（不推荐）。"""
    if last_row is None:
        return f"{sheet}!${col}:${col}"
    return f"{sheet}!${col}${first_row}:${col}${last_row}"


# ---------------------------------------------------------------- 求和 / 计数

def build_sumifs_formula(sum_col: str, cond_map: dict, source_sheet: str = RAW,
                         first_row: int = 3, last_row=None) -> str:
    """
    生成 SUMIFS 求和公式。
    :param sum_col: 求和列字母，例 "F"
    :param cond_map: 条件字典 {条件列字母: 条件单元格引用或字面量}
    """
    parts = [_rng(sum_col, first_row, last_row, source_sheet)]
    for c_col, c_cell in cond_map.items():
        parts.append(_rng(c_col, first_row, last_row, source_sheet))
        parts.append(str(c_cell))
    return "=SUMIFS(%s)" % ",".join(parts)


def build_countifs_formula(cond_map: dict, source_sheet: str = RAW,
                          first_row: int = 3, last_row=None) -> str:
    """生成 COUNTIFS 计数公式（涉及数据笔数 = 该分组原始明细行数）。"""
    parts = []
    for c_col, c_cell in cond_map.items():
        parts.append(_rng(c_col, first_row, last_row, source_sheet))
        parts.append(str(c_cell))
    return "=COUNTIFS(%s)" % ",".join(parts)


def build_sum_all_formula(col: str, source_sheet: str = RAW,
                          first_row: int = 3, last_row=None) -> str:
    """全表求和：SUM(区间)。"""
    return "=SUM(%s)" % _rng(col, first_row, last_row, source_sheet)


def build_counta_formula(col: str, source_sheet: str = RAW,
                         first_row: int = 3, last_row=None) -> str:
    """全表非空计数（总数据笔数）。"""
    return "=COUNTA(%s)" % _rng(col, first_row, last_row, source_sheet)


def build_distinct_all_formula(col: str, source_sheet: str = RAW,
                               first_row: int = 3, last_row=None) -> str:
    """全表去重计数（总合同数）：SUMPRODUCT(1/COUNTIF(区间,区间&""))。"""
    r = _rng(col, first_row, last_row, source_sheet)
    return '=SUMPRODUCT(1/COUNTIF(%s,%s&""))' % (r, r)


def build_distinct_nonblank_all_formula(col: str, source_sheet: str = RAW,
                                        first_row: int = 3, last_row=None) -> str:
    """全表非空去重计数（未回款原因分类数）：空值不计入分类数。"""
    r = _rng(col, first_row, last_row, source_sheet)
    return '=SUMPRODUCT((%s<>"")/COUNTIF(%s,%s&""))' % (r, r, r)


def build_sumproud_distinct_contract_formula(group_col: str, val_cell: str,
                                             contract_col: str, source_sheet: str = RAW,
                                             first_row: int = 3, last_row=None) -> str:
    """
    分组条件下【不重复合同数】。
    口径：该分组取值的原始明细行中，不同合同号的个数（同一合同号在组内多行只计 1 次）。

    ⚠️ **条件侧的列必须带 `&""`**：`COUNTIFS(合同号,合同号&"",维度,维度&"")`。
    COUNTIFS 的「条件」若是**裸单元格引用且为空**，Excel 会按 **0** 处理（而非空白），
    分母遂为 0 → `0/0` → **`#DIV/0!`**。「未回款原因分类」对已回款行为空（本数据集 431 行），
    分组去重计数正踩此坑。

    2026-10-03 修复：重构时漏掉了维度侧的 `&""`（`_recovery/pathA/formulas.py` 原实现
    两侧都有），使「涉及合同数」在 Excel/WPS 打开重算后报错；而注入的缓存值是 pandas
    算的、依然正确，自检只读缓存值 —— 两个缺陷相互掩盖。已新增 R18 对公式文本直接断言。
    """
    c = _rng(contract_col, first_row, last_row, source_sheet)
    d = _rng(group_col, first_row, last_row, source_sheet)
    return '=SUMPRODUCT((%s=%s)/COUNTIFS(%s,%s&"",%s,%s&""))' % (d, val_cell, c, c, d, d)


# ---------------------------------------------------------------- 除法 / 比率

def build_divide_formula(num_cell: str, den_cell: str) -> str:
    """同表内派生比率：分母为 0 时输出空值。=IF(C3=0,"",E3/C3)"""
    return '=IF(%s=0,"",%s/%s)' % (den_cell, num_cell, den_cell)


def build_divide_by_global_formula(num_cell: str, den_col: str, source_sheet: str = RAW,
                                   first_row: int = 3, last_row=None) -> str:
    """本行数值 / 全表某列合计（未回款占比用），分母为全表 SUM。"""
    s = build_sum_all_formula(den_col, source_sheet, first_row, last_row)[1:]
    return '=IF(%s=0,"",%s/%s)' % (s, num_cell, s)


def build_divide_by_columns_diff_formula(num_cell: str, minuend_col: str, subtrahend_col: str,
                                         source_sheet: str = RAW,
                                         first_row: int = 3, last_row=None) -> str:
    """
    本行数值 / (数据源两列合计之差)。用于「未回款原因分类汇总」的占比：

        分母 = 合同金额合计 - 回款合计（即统计总览的"总合同金额 - 总回款合计"），
        两列合计均直接取自「原始数据」表。

    形如：
        =IF((SUM(原始数据!$F$3:$F$1002)-SUM(原始数据!$J$3:$J$1002))=0,"",
            C3/(SUM(原始数据!$F$3:$F$1002)-SUM(原始数据!$J$3:$J$1002)))

    全部引用「原始数据」区间，满足硬性规则"统计值必须动态引用数据源"；
    分母为 0 时输出空值，避免 `#DIV/0!`。
    """
    a = build_sum_all_formula(minuend_col, source_sheet, first_row, last_row)[1:]
    b = build_sum_all_formula(subtrahend_col, source_sheet, first_row, last_row)[1:]
    den = "(%s-%s)" % (a, b)
    return '=IF(%s=0,"",%s/%s)' % (den, num_cell, den)


# ---------------------------------------------------------------- 序号 / 取首条非空

def build_seq_formula(offset: int = 2) -> str:
    """连续序号：数据从第 3 行开始，故 =ROW()-2。"""
    return "=ROW()-%d" % offset


def build_index_first_non_null(group_col: str, group_val_cell: str, fetch_col: str,
                               source_sheet: str = RAW, first_row: int = 3,
                               last_row=None) -> str:
    """
    取"该分组下第一条非空有效值"。
    内层 INDEX(...,0) 使数组在 Excel 中无需 Ctrl+Shift+Enter 即可求值；
    无匹配时 IFERROR 回退为空串（语义即"该分组该字段全为空"）。
    """
    g = _rng(group_col, first_row, last_row, source_sheet)
    v = _rng(fetch_col, first_row, last_row, source_sheet)
    inner = 'INDEX((%s=%s)*(%s<>""),0)' % (g, group_val_cell, v)
    return '=IFERROR(INDEX(%s,MATCH(1,%s,0)),"")' % (v, inner)


def build_textjoin_distinct_formula(group_col: str, group_val_cell: str, target_col: str,
                                    sep: str = "、", source_sheet: str = RAW,
                                    first_row: int = 3, last_row=None) -> str:
    """
    保留 Excel365 动态数组写法（TEXTJOIN+UNIQUE+FILTER）。

    ⚠️ 本引擎**不调用**本函数：组内去重拼接由 pandas 算出文本后直接写入单元格
    （与 analyze.py 的处理一致，登记为"公式化例外"），避免低版本 Excel 出现 #NAME?。
    """
    g = _rng(group_col, first_row, last_row, source_sheet)
    t = _rng(target_col, first_row, last_row, source_sheet)
    return '=TEXTJOIN("%s",TRUE,UNIQUE(FILTER(%s,%s=%s)))' % (sep, t, g, group_val_cell)
