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
import unicodedata

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

from config_validator import validate_sheet_field_kinds


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

def display_width(text) -> int:
    """
    按**东亚宽度标准**计算显示宽度：`unicodedata.east_asian_width` 为 W（宽）/ F（全角）
    的字符计 2，其余（半角字母数字、ASCII 标点等）计 1。

    覆盖范围远大于"仅汉字区间 \u4e00-\u9fff"，还包括全角标点（，。：；「」（））、
    全角字母数字、假名、谚文等——否则「总合同金额（万元）」这类标题会被低估宽度而显示不全。
    """
    if text is None:
        return 0
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in str(text))


def format_display(value, number_format=None) -> str:
    """
    把单元格值渲染成"实际显示出来的样子"，仅用于估算列宽。

    例：903.192 配 `#,##0.000` → `903.192`；0.2543 配 `0.00%` → `25.43%`。
    直接用原始值会漏掉千分位与百分号占用的宽度。
    """
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        nf = (number_format or "General").strip()
        frac = nf.split(".", 1)[1] if "." in nf else ""
        dec = sum(1 for c in frac if c in "0#")
        if "%" in nf:
            return "{:.{d}f}%".format(float(value) * 100, d=dec)
        if nf in ("General", "@", ""):
            return str(value)
        if "," in nf:
            return "{:,.{d}f}".format(float(value), d=dec)
        return "{:.{d}f}".format(float(value), d=dec)
    return str(value)


def auto_fit_column_width(ws: Worksheet, style_cfg: dict = None, display_lookup: dict = None):
    """
    全部列宽按"最终显示出来的内容"自动适配：

    1. 中文等全角字符按 **2** 个字符宽度计（东亚宽度 W/F），半角字符按 1 计；
    2. 公式单元格取注入的**缓存值**（`display_lookup`，形如 {坐标: 值}）来估宽，
       而不是拿公式文本去量——否则长公式会把列宽撑爆；无缓存值时跳过该格；
    3. 列宽 = clamp(max(表头宽度, 各数据行显示宽度) + padding, min, max)，
       上下限与内边距由配置 `style_setting` 的
       `col_width_min` / `col_width_max` / `col_width_padding` 控制；
    4. 跳过合并单元格（大标题横跨全表，不能用来决定单列宽度）。
    """
    style_cfg = style_cfg or {}
    lo = style_cfg.get("col_width_min", 8)
    hi = style_cfg.get("col_width_max", 40)
    pad = style_cfg.get("col_width_padding", 3)

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        widths = []
        for cell in ws[col_letter]:
            if cell.coordinate in ws.merged_cells:
                continue
            coord = cell.coordinate
            if display_lookup and coord in display_lookup:
                text = format_display(display_lookup[coord], cell.number_format)
            else:
                val = cell.value
                if val is None:
                    continue
                if isinstance(val, str) and val.startswith("="):
                    continue          # 无缓存值可用时，不拿公式文本当宽度
                text = format_display(val, cell.number_format)
            widths.append(display_width(text))
        w = (max(widths) + pad) if widths else lo
        ws.column_dimensions[col_letter].width = max(lo, min(hi, float(w)))


def _looks_numeric(value) -> bool:
    """判断值是否“看起来是数值”：数字本身，或能被 float() 解析的字符串。"""
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        t = value.strip()
        if not t:
            return False
        try:
            float(t)
            return True
        except ValueError:
            return False
    return False


def detect_numeric_text_columns(ws: Worksheet, header_row_idx: int, name2idx: dict,
                                text_names: list, min_samples: int = 3,
                                numeric_ratio: float = 0.8) -> list:
    """
    找出「登记为文本字段、但数据几乎全是数值」的列，供告警使用（不阻断生成）。

    这类列通常是**类型登记错了**：本该是数值列（右对齐 + 金额/百分比格式），
    却被放进了 text_fields，于是左对齐且不设数字格式。

    :return: [(表头, 列号, 数值占比), ...]
    """
    found = []
    if not text_names:
        return found
    for name in text_names:
        c_idx = name2idx.get(name)
        if not c_idx:
            continue
        vals = [ws.cell(r, c_idx).value for r in range(header_row_idx + 1, ws.max_row + 1)]
        vals = [v for v in vals if v not in (None, "")]
        if len(vals) < min_samples:
            continue
        hit = sum(1 for v in vals if _looks_numeric(v))
        ratio = hit / float(len(vals))
        if ratio >= numeric_ratio:
            found.append((name, c_idx, ratio))
    return found


# def apply_sheet_format(ws: Worksheet, style_cfg: dict,
#                        num_cols: list, pct_cols: list, date_cols: list, cat_cols: list,
#                        header_row_idx: int = 2, enable_cat_cond_format: bool = False, cat_field_name: str = "",
#                        sheet_title: str = None, total_title_cols: int = 0):
def apply_sheet_format(ws: Worksheet, style_cfg: dict,
                       num_cols: list, pct_cols: list, date_cols: list, cat_cols: list,
                       header_row_idx: int = 2, enable_cat_cond_format: bool = False, cat_field_name: str = "",
                       sheet_title: str = None, total_title_cols: int = 0, int_cols: list = None,
                       text_cols: list = None,
                       set_number_format: bool = True, display_values: dict = None):
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
    :param text_cols:文本字段名列表（显式左对齐）；所有列都必须在类型清单中登记
    :param display_values:公式格的缓存值 {坐标: 值}；列宽按"显示内容"估算时使用，
                          不传则公式格不参与列宽计算（只按表头与文本格估宽）
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

    # 字段类型清单：每一列都必须显式登记；有未登记的表头直接报错，
    # 避免新增列被静默地兜底成文本（左对齐、不设数字格式）。
    int_cols = int_cols or []
    text_cols = text_cols or []
    validate_sheet_field_kinds(
        ws.title, list(name2idx.keys()),
        {"num_fields": num_cols, "int_fields": int_cols, "pct_fields": pct_cols,
         "date_fields": date_cols, "cat_fields": cat_cols, "text_fields": text_cols},
    )

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
    int_cols = int_cols or []
    # 遍历数据行：斑马纹（奇数数据行上底色，偶数行不填充）+ 浅灰水平边框
    for r in range(header_row_idx + 1, max_row + 1):
        striped = ((r - header_row_idx) % 2 == 1)
        row_fill = odd_fill if striped else PatternFill(fill_type=None)
        for cell in ws[r]:
            cell.fill = row_fill
            cell.border = h_border
        # 按字段类型设置对齐与数字格式
        for col_name in name2idx:
            c_idx = name2idx[col_name]
            cell = ws.cell(row=r, column=c_idx)
            if col_name in num_cols:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if set_number_format:
                    cell.number_format = style_cfg["amount_number_format"]
            elif col_name in int_cols:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if set_number_format:
                    cell.number_format = style_cfg.get("int_number_format", "0")
            elif col_name in pct_cols:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if set_number_format:
                    cell.number_format = style_cfg["percent_number_format"]
            elif col_name in date_cols:
                # 日期字段居中；维持源文本录入格式（默认 @，不得转成日期序列值）
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if set_number_format:
                    cell.number_format = style_cfg.get("date_number_format", "@")
            elif col_name in cat_cols:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in text_cols:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # 任务2：未回款原因分类差异化条件格式
    if enable_cat_cond_format and cat_field_name in name2idx:
        cat_col = name2idx[cat_field_name]
        add_category_conditional_format(ws, cat_col, start_row=header_row_idx + 1)

    # 告警：被登记为文本字段、但数据几乎全是数值的列（很可能是类型登记错了）。
    # 只提示不阻断——真正的"未登记列"已在上面 validate_sheet_field_kinds 直接报错。
    for _name, _col, _ratio in detect_numeric_text_columns(ws, header_row_idx, name2idx, text_cols):
        print("⚠️ [对齐告警]「%s」%s 列被登记为文本字段，但 %.0f%% 的数据是数值；"
              "如确为数值请移入 num_fields（右对齐 + 数字格式）"
              % (ws.title, get_column_letter(_col), _ratio * 100))

    auto_fit_column_width(ws, style_cfg, display_values)