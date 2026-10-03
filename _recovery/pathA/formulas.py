# -*- coding: utf-8 -*-
"""
公式构造器。

硬性规则：所有统计、计算单元格必须是动态公式，禁止写入 Python 算出的常量。
本模块只负责"拼公式字符串"；公式对应的期望值由 analyze.py 独立用 Python 计算，
写入 cached_values.py 描述的缓存注入流程，两者互不复用（避免自证）。

约定：
  - 数据源统一为「原始数据」工作表的固定区间（首末行由 analyze.py 探测后传入）
  - 表名一律加单引号，兼容中文表名
  - 区间一律绝对引用（$），保证下拉填充与评审抽查时引用不漂移
"""

RAW = "'原始数据'"


def _rng(col_letter, first, last):
    """原始数据表的绝对列区间，如 '原始数据'!$F$3:$F$1002"""
    return "%s!$%s$%d:$%s$%d" % (RAW, col_letter, first, col_letter, last)


# ---------------------------------------------------------------- 求和 / 计数

def sumif(dim_col, key_ref, val_col, first, last):
    """按维度求和：SUMIF(维度区间, 本行维度值, 数值区间)"""
    return "=SUMIF(%s,%s,%s)" % (_rng(dim_col, first, last), key_ref,
                                 _rng(val_col, first, last))


def countif(dim_col, key_ref, first, last):
    """按维度计笔数：COUNTIF(维度区间, 本行维度值)"""
    return "=COUNTIF(%s,%s)" % (_rng(dim_col, first, last), key_ref)


def sum_all(val_col, first, last):
    """全表求和：SUM(数值区间)"""
    return "=SUM(%s)" % (_rng(val_col, first, last))


def count_all(val_col, first, last):
    """全表计非空笔数：COUNTA(区间)"""
    return "=COUNTA(%s)" % (_rng(val_col, first, last))


def distinct_count_all(code_col, first, last):
    """全表去重计数（总合同数）：SUMPRODUCT(1/COUNTIF(合同号区间, 合同号区间))"""
    r = _rng(code_col, first, last)
    return "=SUMPRODUCT(1/COUNTIF(%s,%s&\"\"))" % (r, r)


def distinct_count_dim(dim_col, key_ref, code_col, first, last):
    """
    分组内不同合同号个数（合同数 / 涉及合同数）。

    口径已用附件2-示例逐格验证：不是明细行数，也不是合并表行数，
    而是"该维度取值的原始明细行中，不同合同号的个数"。
    同一合同号若在组内出现多行只计 1 次；跨组则各组分别计 1 次。

    公式：SUMPRODUCT( (维度=本行维度) / COUNTIFS(合同号,合同号, 维度,维度) )
    分母对每一行都 >=1（每行至少匹配自身），不存在除零。
    """
    d = _rng(dim_col, first, last)
    c = _rng(code_col, first, last)
    return "=SUMPRODUCT((%s=%s)/COUNTIFS(%s,%s&\"\",%s,%s&\"\"))" % (d, key_ref, c, c, d, d)


def distinct_count_nonblank_all(val_col, first, last):
    """
    全表某字段的非空去重个数（统计总览「未回款原因分类数」）。
    空值不计入分类数：SUMPRODUCT((区间<>"")/COUNTIF(区间,区间))
    """
    r = _rng(val_col, first, last)
    return "=SUMPRODUCT((%s<>\"\")/COUNTIF(%s,%s&\"\"))" % (r, r, r)


# ---------------------------------------------------------------- 比率

def ratio(num_ref, den_ref):
    """单元格相除，分母为 0 时输出空值（不把未知包装成 0）"""
    return "=IF(%s=0,\"\",%s/%s)" % (den_ref, num_ref, den_ref)


def ratio_total(num_ref, val_col, first, last):
    """本组数值 / 全局合计（未回款占比用），分母为全表 SUM"""
    return "=IF(%s=0,\"\",%s/%s)" % (sum_all(val_col, first, last)[1:], num_ref,
                                    sum_all(val_col, first, last)[1:])


# ---------------------------------------------------------------- 取首条非空值

def first_value(code_col, key_ref, val_col, first, last):
    """
    取"该合同号分组下第一条非空有效值"。

    实现用 INDEX + MATCH(1, INDEX((条件1)*(条件2),0), 0)：
    内层 INDEX(...,0) 让数组在 Excel 中无需 Ctrl+Shift+Enter 即可求值，
    LibreOffice 同样支持。无匹配时 IFERROR 回退为空串（语义即"无非空值"）。
    """
    c = _rng(code_col, first, last)
    v = _rng(val_col, first, last)
    inner = "INDEX((%s=%s)*(%s<>\"\"),0)" % (c, key_ref, v)
    return "=IFERROR(INDEX(%s,MATCH(1,%s,0)),\"\")" % (v, inner)


def first_value_simple(code_col, key_ref, val_col, first, last):
    """
    降级实现：取该合同号第一次出现那一行的值。
    仅当该字段在源数据中完全无空值时，与 first_value 等价（本数据集即属此情况）。
    """
    c = _rng(code_col, first, last)
    v = _rng(val_col, first, last)
    return "=IFERROR(INDEX(%s,MATCH(%s,%s,0)),\"\")" % (v, key_ref, c)


# ---------------------------------------------------------------- 行号序号

def seq_by_row(offset=2):
    """全新连续序号：=ROW()-2（数据从第 3 行开始）"""
    return "=ROW()-%d" % offset
