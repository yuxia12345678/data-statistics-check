# -*- coding: utf-8 -*-
"""
Excel动态公式生成器 formula_builder
核心：全部输出Excel原生公式，不硬编码业务数值，源表为【原始数据】工作表
支持：SUMIFS、COUNTIFS、SUMPRODUCT(去重计数)、IFERROR防除零、TEXTJOIN、INDEX+MATCH
说明：所有公式引用原始数据工作表，数据更新后，打开Excel自动重算；
本模块只输出公式字符串，不做实际内存计算，实现动态联动。
"""


def build_sumifs_formula(sum_col: str, cond_map: dict, source_sheet: str = "原始数据") -> str:
    """
    生成SUMIFS求和公式
    :param sum_col: 需要求和的列字母，例 "K"
    :param cond_map: 条件字典 {条件列字母:条件单元格引用}
    :param source_sheet: 数据源工作表名称
    :return: Excel公式字符串
    """
    cond_parts = []
    for c_col, c_cell in cond_map.items():
        cond_parts.append(f"{source_sheet}!${c_col}:${c_col}")
        cond_parts.append(c_cell)
    inner_args = ",".join([f"{source_sheet}!${sum_col}:${sum_col}"] + cond_parts)
    return f'=SUMIFS({inner_args})'


def build_countifs_formula(cond_map: dict, source_sheet: str = "原始数据") -> str:
    cond_parts = []
    for c_col, c_cell in cond_map.items():
        cond_parts.append(f"{source_sheet}!${c_col}:${c_col}")
        cond_parts.append(c_cell)
    inner_args = ",".join(cond_parts)
    return f'=COUNTIFS({inner_args})'



def build_countifs_formula(cond_map: dict, source_sheet: str = "原始数据") -> str:
    """生成COUNTIFS计数公式"""
    cond_parts = []
    for c_col, c_cell in cond_map.items():
        cond_parts.append(f"{source_sheet}!${c_col}:${c_col}")
        cond_parts.append(c_cell)
    inner_args = ",".join(cond_parts)
    return f'=COUNTIFS({inner_args})'


# def build_sumproud_distinct_contract_formula(group_col: str, val_cell: str,
#                                              contract_col: str, source_sheet: str = "原始数据") -> str:
#     """
#     SUMPRODUCT实现：分组条件下【不重复合同数】统计
#     业务：任务2涉及合同数：满足条件的不重复合同号数量
#     """
#     formula = (
#         f'=SUMPRODUCT(1/COUNTIF({source_sheet}!${contract_col}:${contract_col},'
#         f'{source_sheet}!${contract_col}:${contract_col})*'
#         f'({source_sheet}!${group_col}:${group_col}={val_cell}))'
#     )
#     return formula

# def build_sumproud_distinct_contract_formula(group_col: str, val_cell: str,
#                                              contract_col: str, source_sheet: str = "原始数据") -> str:
#     """
#     SUMPRODUCT实现：分组条件下【不重复合同数】统计
#     修复：使用 COUNTIFS + 非空判断，避免原始数据存在空白单元格导致 #DIV/0! 报错
#     """
#     formula = (
#         f'=SUMPRODUCT(({source_sheet}!${group_col}:${group_col}={val_cell})*'
#         f'({source_sheet}!${contract_col}:${contract_col}<>"")/'
#         f'COUNTIFS({source_sheet}!${contract_col}:${contract_col},{source_sheet}!${contract_col}:${contract_col}&"",'
#         f'{source_sheet}!${group_col}:${group_col},{source_sheet}!${group_col}:${group_col}))'
#     )
#     return formula
def build_sumproud_distinct_contract_formula(group_col: str, val_cell: str,
                                             contract_col: str, source_sheet: str = "原始数据") -> str:
    """
    SUMPRODUCT实现：分组条件下【不重复合同数】统计
    修复：使用 IF(COUNTIFS(...)=0, 1, COUNTIFS(...)) 避免空白行导致的 #DIV/0! 错误，同时不丢失真实合同数
    """
    countifs_expr = (
        f'COUNTIFS({source_sheet}!${contract_col}:${contract_col},{source_sheet}!${contract_col}:${contract_col}&"",'
        f'{source_sheet}!${group_col}:${group_col},{source_sheet}!${group_col}:${group_col})'
    )
    formula = (
        f'=SUMPRODUCT(({source_sheet}!${group_col}:${group_col}={val_cell})*'
        f'({source_sheet}!${contract_col}:${contract_col}<>"")/'
        f'IF({countifs_expr}=0, 1, {countifs_expr}))'
    )
    return formula

def build_divide_formula(num_cell: str, den_cell: str) -> str:
    """
    生成除法公式，IFERROR捕获分母为0，返回空；用于占比、回款率
    :param num_cell:分子单元格
    :param den_cell:分母单元格
    :return: =IFERROR(A1/B1,"")
    """
    return f'=IFERROR({num_cell}/{den_cell},"")'


def build_textjoin_distinct_formula(group_col: str, group_val_cell: str,
                                    target_col: str, sep: str = "；",
                                    source_sheet: str = "原始数据") -> str:
    """
    TEXTJOIN 去重拼接：同一合同号未回款原因分类拼接
    需要Excel365支持动态数组
    """
    formula = (
        f'=TEXTJOIN("{sep}",TRUE,UNIQUE(FILTER({source_sheet}!${target_col}:${target_col},'
        f'{source_sheet}!${group_col}:${group_col}={group_val_cell})))'
    )
    return formula


def build_index_first_non_null(group_col: str, group_val_cell: str, fetch_col: str, source_sheet: str = "原始数据") -> str:
    # 简化为普通的 MATCH 查找，并加上 IFERROR 防错
    formula = (
        f'=IFERROR(INDEX({source_sheet}!${fetch_col}:${fetch_col},'
        f'MATCH({group_val_cell},{source_sheet}!${group_col}:${group_col},0)),"")'
    )
    return formula
