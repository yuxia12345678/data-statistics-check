# -*- coding: utf-8 -*-
"""
Excel统一美化渲染模块 excel_styler
严格遵循题目可视化规范：
1、表头低饱和度蓝色背景，白色加粗居中；
2、斑马纹交替底色，仅保留浅灰色水平边框；
3、对齐：文本左对齐；数值右对齐；日期、分类居中；
4、自动列宽，中文全角字符权重2；
5、冻结表头首行；开启auto_filter筛选；
6、数字格式：金额千分位3位小数；百分比2位小数；
7、支持任务2：未回款原因分类差异化柔和背景条件格式
8、工作表第1行合并大标题
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter


def build_excel_style_from_config(style_cfg: dict):
    """读取配置构建可复用openpyxl样式对象"""
    header_fill = PatternFill(start_color=style_cfg["header_bg_color"],
                              end_color=style_cfg["header_bg_color"], fill_type="solid")
    header_font = Font(bold=True, color=style_cfg["header_font_color"])
    odd_row_fill = PatternFill(start_color=style_cfg["odd_row_bg"], end_color=style_cfg["odd_row_bg"], fill_type="solid")
    even_row_fill = PatternFill(start_color=style_cfg["even_row_bg"], end_color=style_cfg["even_row_bg"], fill_type="solid")
    thin_h_line = Side(style="thin", color=style_cfg["horizontal_border_color"])
    h_border = Border(top=None, bottom=thin_h_line, left=None, right=None)

    # 大标题样式
    title_fill = PatternFill(start_color=style_cfg["title_bg_color"],
                             end_color=style_cfg["title_bg_color"], fill_type="solid")
    title_font = Font(bold=True, color=style_cfg["title_font_color"], size=style_cfg["title_font_size"])

    return {
        "header_fill": header_fill,
        "header_font": header_font,
        "odd_fill": odd_row_fill,
        "even_fill": even_row_fill,
        "h_border": h_border,
        "title_fill": title_fill,
        "title_font": title_font
    }


def set_sheet_title_row(ws: Worksheet, title_text: str, total_col_count: int):
    """
    设置工作表第1行合并大标题
    :param ws: 工作表对象
    :param title_text: 标题显示文本
    :param total_col_count: 需要合并的总列数量
    """
    from openpyxl.worksheet.merge import MergeCell
    end_letter = get_column_letter(total_col_count)
    merge_range = f"A1:{end_letter}1"
    ws.merge_cells(merge_range)
    cell = ws["A1"]
    cell.value = title_text
    # 样式在build_excel_style_from_config获取，调用方传入
    return merge_range


def add_category_conditional_format(ws: Worksheet, cat_col_index: int, start_row: int):
    """
    任务2专属：未回款原因分类列设置差异化柔和背景条件格式
    :param ws:工作表对象
    :param cat_col_index:分类列序号，从1开始
    :param start_row:数据起始行（跳过表头）
    """
    soft_color_list = ["FFF2CC", "E2EFDA", "E7E6E6", "DDEBF7", "FCE4D6", "F2F2F2", "E5E7EB"]
    max_row = ws.max_row
    range_str = f"{ws.cell(start_row, cat_col_index).coordinate}:{ws.cell(max_row, cat_col_index).coordinate}"
    used_values = set()
    for r in range(start_row, max_row + 1):
        cell = ws.cell(row=r, column=cat_col_index)
        val = cell.value
        if val is None or val in used_values:
            continue
        used_values.add(val)
        pick_color = soft_color_list[len(used_values) % len(soft_color_list)]
        fill = PatternFill(start_color=pick_color, end_color=pick_color, fill_type="solid")
        rule = FormulaRule(formula=[f'${ws.cell(r, cat_col_index).column_letter}{r}="{val}"'],
                           fill=fill, stopIfTrue=True)
        ws.conditional_formatting.add(range_str, rule)

# def auto_fit_column_width(ws: Worksheet):
#     """
#     自适应列宽；中文全角字符权重2，英文数字权重1；限制最小8，最大45；跳过合并单元格
#     """
#     for col_idx, col in enumerate(ws.columns, start=1):
#         max_len = 0
#         # 使用 openpyxl.utils 的 get_column_letter，根据列的索引获取列字母，而不是依赖单元格属性
#         col_letter = get_column_letter(col_idx)
#         for cell in col:
#             # 跳过合并单元格，避免出现其他属性错误
#             if cell.coordinate in ws.merged_cells:
#                 continue
#             if cell.value:
#                 val = str(cell.value)
#                 length = sum([2 if '\\u4e00' <= c <= '\\u9fff' else 1 for c in val])
#                 max_len = max(max_len, length)
#         final_w = max(8, min(max_len + 3, 45))
#         ws.column_dimensions[col_letter].width = final_w

def auto_fit_column_width(ws: Worksheet):
    """
    自适应列宽；中文全角字符权重2，英文数字权重1；限制最小12，最大35；跳过合并单元格；
    重点优化：跳过公式字符串，防止长公式撑爆列宽。
    """
    for col_idx, col in enumerate(ws.columns, start=1):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        has_formula = False      
        for cell in col:
            if cell.coordinate in ws.merged_cells:
                continue
            if cell.value is not None:
                val = str(cell.value)              
                # 核心修复：遇到公式（以=开头），不计算字符长度，但标记该列有公式
                if val.startswith("="):
                    has_formula = True
                    continue                 
                length = sum([2 if '\\u4e00' <= c <= '\\u9fff' else 1 for c in val])
                max_len = max(max_len, length)       
        # 针对包含公式的列，由于没有具体的数值长度，给予一个合理的预估宽度（比如 12）
        if has_formula and max_len < 12:
            max_len = 12
            
        # 调整列宽的上下限，让表格更美观：最小 12，最大 35
        final_w = max(12, min(max_len + 2, 35))
        ws.column_dimensions[col_letter].width = final_w


# def apply_sheet_format(ws: Worksheet, style_cfg: dict,
#                        num_cols: list, pct_cols: list, date_cols: list, cat_cols: list,
#                        header_row_idx: int = 2, enable_cat_cond_format: bool = False, cat_field_name: str = "",
#                        sheet_title: str = None, total_title_cols: int = 0):
def apply_sheet_format(ws: Worksheet, style_cfg: dict,
                       num_cols: list, pct_cols: list, date_cols: list, cat_cols: list,
                       header_row_idx: int = 2, enable_cat_cond_format: bool = False, cat_field_name: str = "",
                       sheet_title: str = None, total_title_cols: int = 0, int_cols: list = None):  # <--- 增加int_cols参数
    """
    对工作表执行全套美化
    :param ws: worksheet对象
    :param style_cfg:样式配置字典
    :param num_cols:金额字段名列表
    :param pct_cols:百分比字段名列表
    :param date_cols:日期字段列表
    :param cat_cols:分类字段列表
    :param header_row_idx:表头行号（新增大标题后，表头固定为第2行）
    :param enable_cat_cond_format:是否开启分类差异化条件格式（任务2打开）
    :param cat_field_name:条件格式绑定的分类字段名称
    :param sheet_title:大标题文本，不为空则设置A1合并标题
    :param total_title_cols:标题合并总列数
    """
    style_dict = build_excel_style_from_config(style_cfg)
    header_fill = style_dict["header_fill"]
    header_font = style_dict["header_font"]
    odd_fill = style_dict["odd_fill"]
    even_fill = style_dict["even_fill"]
    h_border = style_dict["h_border"]
    title_fill = style_dict["title_fill"]
    title_font = style_dict["title_font"]

    # 设置第1行合并大标题
    if sheet_title is not None and total_title_cols > 0:
        set_sheet_title_row(ws, sheet_title, total_title_cols)
        ws["A1"].fill = title_fill
        ws["A1"].font = title_font
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # 冻结窗格：大标题1行 + 表头1行，数据从第3行开始
    ws.freeze_panes = f"A{header_row_idx + 1}"
    header_names_raw = [cell.value for cell in ws[header_row_idx]]
    name2idx = {}
    for idx, n in enumerate(header_names_raw):
        if n is not None:
            name2idx[n] = idx + 1

    # 设置表头样式：蓝色底色，白色加粗，水平居中
    # for cell in ws[header_row_idx]:
    #     cell.fill = header_fill
    #     cell.font = header_font
    #     cell.alignment = Alignment(horizontal="center", vertical="center")

    # # 开启表头筛选，筛选区域从表头行开始
    # #ws.auto_filter.ref = ws[f"{header_row_idx}:{ws.max_row}"]
    # # 开启表头筛选，筛选区域从表头行开始
    # # 修复：必须使用标准的 Excel 范围字符串（例如 "A2:E10"），不能使用 ws[...] 切片
    # max_col_letter = get_column_letter(ws.max_column)
    # ws.auto_filter.ref = f"A{header_row_idx}:{max_col_letter}{ws.max_row}"
    
    # 修复：使用传入的 total_title_cols 或有效列数来限制表头和筛选范围
    if total_title_cols > 0:
        # 仅对真实存在的列应用表头样式
        for col_idx in range(1, total_title_cols + 1):
            cell = ws.cell(row=header_row_idx, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        # 筛选范围严格限制在有效列内
        max_col_letter = get_column_letter(total_title_cols)
        ws.auto_filter.ref = f"A{header_row_idx}:{max_col_letter}{ws.max_row}"
    else:
        # 兜底逻辑：如果未指定列数，则按原有方式执行
        for cell in ws[header_row_idx]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        max_col_letter = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A{header_row_idx}:{max_col_letter}{ws.max_row}"
    
    max_row = ws.max_row
    # 遍历数据行，斑马纹底色 + 水平边框
    for r in range(header_row_idx + 1, max_row + 1):
        row_cells = ws[r]
        row_fill = odd_fill if (r % 2 == 1) else even_fill
        for cell in row_cells:
            cell.fill = row_fill
            cell.border = h_border
        # 按字段类型设置对齐与数字格式
        int_cols = int_cols or []
        for col_name in name2idx:
            c_idx = name2idx[col_name]
            cell = ws.cell(row=r, column=c_idx)
            if col_name in num_cols:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = style_cfg["amount_number_format"]
            elif col_name in int_cols:                                     # <--- 新增整数判定
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = style_cfg.get("int_number_format", "0")
            elif col_name in pct_cols:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = style_cfg["percent_number_format"]
            elif col_name in date_cols or col_name in cat_cols:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # 任务2：未回款原因分类差异化条件格式
    if enable_cat_cond_format and cat_field_name in name2idx:
        cat_col = name2idx[cat_field_name]
        add_category_conditional_format(ws, cat_col, start_row=header_row_idx + 1)

    auto_fit_column_width(ws)