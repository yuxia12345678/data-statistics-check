# -*- coding: utf-8 -*-
# ============================================================
# 合同数据统计核对 SKILL 主程序
# ============================================================
# 功能：读取《合同开票及回款核对表.xlsx》，自动完成：
#   1. 数据清洗合并（同合同号多行合并为一行）
#   2. 未回款原因分类汇总
#   3. 多维度统计分析（区域/部门/账龄/客户分类）
#   4. 统计总览
#   5. 全工作表统一美化
# 输出：包含8张工作表的完整结果Excel文件
# ============================================================

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import OrderedDict
#后面调用skill的时候取消掉
from pathlib import Path

# ============================================================
# 1. 读取源数据
# ============================================================
# 当前脚本所在目录
BASE_DIR = Path(__file__).resolve().parent

# 拼出同目录下的文件
src_path = BASE_DIR / '附件1-合同开票及回款核对表.xlsx'
output_path = BASE_DIR / '合同开票及回款核对表分析结果_测什么都对.xlsx'

wb_src = openpyxl.load_workbook(src_path, read_only=True, data_only=True)
ws_src = wb_src.active

# 读取表头（第2行）
headers = []
for cell in ws_src[2]:
    headers.append(cell.value if cell.value is not None else '')
# 去掉末尾空列
while headers and headers[-1] == '':
    headers.pop()

# 全量读取数据（第3行起）
data = []
for row in ws_src.iter_rows(min_row=3, values_only=True):
    row_data = {}
    for i, h in enumerate(headers):
        if i < len(row):
            row_data[h] = row[i]
        else:
            row_data[h] = None
    data.append(row_data)
wb_src.close()

# ============================================================
# 2. 样式定义
# ============================================================

# 大标题样式
TITLE_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="FFFFFF")
TITLE_ALIGN = Alignment(horizontal="center", vertical="center")

# 表头样式
HEADER_FILL = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
HEADER_FONT = Font(name='微软雅黑', bold=True, color="FFFFFF", size=11)
HEADER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

ZEBRA_ODD = PatternFill(start_color="EBF1F8", end_color="EBF1F8", fill_type="solid")
ZEBRA_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
# 边框
LIGHT_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

TEXT_ALIGN = Alignment(horizontal='left', vertical='center')
NUM_ALIGN = Alignment(horizontal='right', vertical='center')
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')

numeric_cols = ['合同金额', '未开票金额', '开票金额', '回款合计', '开票未回款', '合并开票金额']
center_cols = ['合同号', '区域', '客户分类', '部门', '责任人', '账龄', '未回款原因分类']
ratio_cols = ['整体回款率', '未回款占比']
serialNum_col = '序号'


def style_title_and_header(ws, title, headers, title_row=1, header_row=2):
    """
    ws:       工作表对象
    title:    第1行的大标题文本
    headers:  第2行的表头列表，如 ['序号','区域',...]
    """
    ncols = len(headers)

    # ---------- 第 1 行：大标题 ----------
    ws.cell(row=title_row, column=1, value=title)
    # 合并 A1 到最后一列
    ws.merge_cells(
        start_row=title_row, start_column=1,
        end_row=title_row, end_column=ncols
    )
    title_cell = ws.cell(row=title_row, column=1)
    title_cell.fill = TITLE_FILL
    title_cell.font = TITLE_FONT
    title_cell.alignment = TITLE_ALIGN
    # 想给标题也加边框，就遍历合并区每个格子
    for c in range(1, ncols + 1):
        ws.cell(row=title_row, column=c).border = LIGHT_BORDER

    # 设置标题行高
    ws.row_dimensions[title_row].height = 24

    # ---------- 第 2 行：表头 ----------
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=text)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = LIGHT_BORDER

    ws.row_dimensions[header_row].height = 20

def style_cell(cell, col_name, value, row_idx, is_header=False):
    """统一设置单元格样式"""
    cell.border = LIGHT_BORDER
    if is_header:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        return
    cell.fill = ZEBRA_ODD if row_idx % 2 == 0 else ZEBRA_EVEN
    if col_name in numeric_cols:
        cell.number_format = '#,##0.000'
        cell.alignment = NUM_ALIGN
    elif col_name in ratio_cols:
        cell.number_format = '0.00%'
        cell.alignment = NUM_ALIGN
    elif col_name in center_cols:
        cell.alignment = CENTER_ALIGN
    elif col_name == serialNum_col:
        cell.alignment = NUM_ALIGN
    else:
        cell.alignment = TEXT_ALIGN


def auto_width(ws, sample_rows=50):
    """自动列宽"""
    for ci in range(1, ws.max_column + 1):
        ml = len(str(ws.cell(1, ci).value or ''))
        for r in range(2, min(ws.max_row + 1, sample_rows + 1)):
            v = ws.cell(r, ci).value
            if v:
                ml = max(ml, min(len(str(v)), 30))
        ws.column_dimensions[get_column_letter(ci)].width = ml + 4


# ============================================================
# 3. 原始数据工作表（原样复制，空值替换为"未填写"）
# ============================================================
wb_out = openpyxl.Workbook()
ws = wb_out.active
ws.title = '原始数据'

headers1 = ["序号","区域","客户分类","合同号","开票日期","合同金额","合并开票金额",
            "未开票金额","开票金额","回款合计","开票未回款","部门","责任人","账龄","未回款原因分类"]
style_title_and_header(ws, "原始数据明细", headers1)


for ri, rd in enumerate(data, 3):
    for ci, h in enumerate(headers, 1):
        val = rd.get(h)
        if val is None or (isinstance(val, str) and val.strip() == ''):
            val = '未填写'
        style_cell(ws.cell(ri, ci, val), h, val, ri)

ws.freeze_panes = 'A2'
auto_width(ws)

# ============================================================
# 4. 同合同号合并汇总
#    - 数值字段：按合同号分组求和
#    - 文本字段：保留第一条非空值
#    - 未回款原因：去重拼接
#    - 序号：重新生成连续序号，替换原始序号列
# ============================================================
ws2 = wb_out.create_sheet('同合同号合并汇总')

groups = OrderedDict()
for rd in data:
    cid = rd.get('合同号')
    if cid is None:
        cid = '未填写'
    else:
        cid = str(cid).strip()
    groups.setdefault(cid, []).append(rd)

# 表头
headers2 = ["序号","合同号","区域","客户分类","开票日期","合同金额","未开票金额",
            "开票金额","回款合计","开票未回款","部门","责任人","账龄","未回款原因分类"]
style_title_and_header(ws2, "同一合同号数据合并汇总", headers2)

ri = 3
seq = 1
for cid, rows in groups.items():
    # 新序号
    style_cell(ws2.cell(ri, 1, seq), '序号', seq, ri)

    # 原始列映射：跳过原始"序号"列
    for ci, h in enumerate(headers):
        if h == '序号':
            continue
        vals = [r.get(h) for r in rows]
        new_ci = ci + 1  # 列号+1（因跳过序号列）

        if h in numeric_cols:
            sv = sum(v for v in vals if isinstance(v, (int, float)))
            if sv == 0:
                sv = None
            style_cell(ws2.cell(ri, new_ci, sv), h, sv, ri)
        elif h == '未回款原因分类':
            ne = [str(v) for v in vals if v is not None and str(v).strip() != '' and str(v) != '未填写']
            uv = list(OrderedDict.fromkeys(ne))
            style_cell(ws2.cell(ri, new_ci, '; '.join(uv) if uv else None), h, None, ri)
        else:
            first = next((v for v in vals if v is not None and str(v).strip() != ''), None)
            style_cell(ws2.cell(ri, new_ci, first), h, first, ri)

    ri += 1
    seq += 1

ws2.freeze_panes = 'B2'
auto_width(ws2)

# ============================================================
# 5. 未回款原因分类汇总
#    - 每类未回款原因：开票未回款金额、涉及合同数、数据笔数、占比
#    - 按开票未回款金额降序排列
# ============================================================
ws3 = wb_out.create_sheet('未回款原因分类汇总')

headers3 = ["序号","未回款原因分类","开票未回款","涉及合同数","涉及笔数","占比"]
style_title_and_header(ws3, "未回款原因分类汇总", headers3)

rg = OrderedDict()
for rd in data:
    reason = rd.get('未回款原因分类')
    if reason is None or str(reason).strip() == '':
        reason = '未填写'
    else:
        reason = str(reason).strip()
    if reason not in rg:
        rg[reason] = {'rows': [], 'contracts': set()}
    rg[reason]['rows'].append(rd)
    rg[reason]['contracts'].add(rd.get('合同号'))

total_kpwr = sum(r.get('开票未回款', 0) for r in data if isinstance(r.get('开票未回款'), (int, float)))

rr = []
for reason, info in rg.items():
    s = sum(r.get('开票未回款', 0) for r in info['rows'] if isinstance(r.get('开票未回款'), (int, float)))
    pct = s / total_kpwr if total_kpwr > 0 else 0
    rr.append({
        '未回款原因分类': reason,
        '开票未回款': s,
        '涉及合同数': len(info['contracts']),
        '涉及笔数': len(info['rows']),
        '占比': pct
    })
rr.sort(key=lambda x: x['开票未回款'], reverse=True)

ri = 3
seq = 1
for item in rr:
    # 新序号
    style_cell(ws3.cell(ri, 1, seq), '序号', seq, ri)
    for ci, h in enumerate(headers3, 1):
        if h == '序号':
            continue
        # 列号+1（因跳过序号列）
        style_cell(ws3.cell(ri, ci, item[h]), h, item[h], ri)
    ri += 1
    seq += 1

ws3.freeze_panes = 'A2'
auto_width(ws3)

# ============================================================
# 6. 多维度统计分析（区域/部门/账龄/客户分类）
#    统一7项指标：总合同金额、总开票金额、总回款合计、总开票未回款、
#    合同数、整体回款率、未回款占比
#    按开票未回款金额降序排列
# ============================================================
#dim_h = ['维度', '总合同金额', '总开票金额', '总回款合计', '总开票未回款', '合同数', '整体回款率', '未回款占比']

dim_h4 = ['序号', '维度', '总合同金额', '总开票金额', '总回款合计', '总开票未回款', '合同数', '整体回款率', '未回款占比']

def write_dim(wb, name, col):
    ws = wb.create_sheet(name)
    dim_h4[1] = col
    if name == '区域统计':                
        style_title_and_header(ws, "区域维度统计分析", dim_h4)
    elif name == '部门统计':      
        style_title_and_header(ws, "部门维度统计分析", dim_h4)
    elif name == '账龄统计':
        style_title_and_header(ws, "账龄维度统计分析", dim_h4)
    elif name == '客户分类统计':
        style_title_and_header(ws, "客户分类维度统计分析", dim_h4)
#    style_cell(ws.cell(1, ci, h), h, None, 2, is_header=True)

    dg = OrderedDict()
    for rd in data:
        dv = rd.get(col)
        if dv is None or str(dv).strip() == '':
            dv = '未填写'
        else:
            dv = str(dv).strip()
        if dv not in dg:
            dg[dv] = {'rows': [], 'contracts': set()}
        dg[dv]['rows'].append(rd)
        dg[dv]['contracts'].add(rd.get('合同号'))

    dr = []
    for dv, info in dg.items():
        s_ht = sum(r.get('合同金额', 0) for r in info['rows']
                   if isinstance(r.get('合同金额'), (int, float)))
        s_kp = sum(r.get('开票金额', 0) for r in info['rows']
                   if isinstance(r.get('开票金额'), (int, float)))
        s_hk = sum(r.get('回款合计', 0) for r in info['rows']
                   if isinstance(r.get('回款合计'), (int, float)))
        s_kpwr = sum(r.get('开票未回款', 0) for r in info['rows']
                    if isinstance(r.get('开票未回款'), (int, float)))
        nc = len(info['contracts'])
        rk = s_hk / s_kp if s_kp > 0 else 0
        wz = s_kpwr / total_kpwr if total_kpwr > 0 else 0
        dr.append({'序号': None,
            '维度': dv, '总合同金额': s_ht, '总开票金额': s_kp,
            '总回款合计': s_hk, '总开票未回款': s_kpwr,
            '合同数': nc, '整体回款率': rk, '未回款占比': wz
        })
    dr.sort(key=lambda x: x['总开票未回款'], reverse=True)

    ri = 3
    seq = 1
    for item in dr:
        # 新序号
        style_cell(ws.cell(ri, 1, seq), '序号', seq, ri)
        for ci, h in enumerate(dim_h4, 1):
            dim_h4[1] = '维度'
            if h == '序号':
                continue   # 列号+1（因跳过序号列）
            style_cell(ws.cell(ri, ci, item[h]), h, item[h], ri)
        ri += 1
        seq += 1

    ws.freeze_panes = 'A2'
    auto_width(ws)
    return len(dr)


write_dim(wb_out, '区域统计', '区域')
write_dim(wb_out, '部门统计', '部门')
write_dim(wb_out, '账龄统计', '账龄')
write_dim(wb_out, '客户分类统计', '客户分类')

# ============================================================
# 7. 统计总览（全局核心指标）
# ============================================================
ws8 = wb_out.create_sheet('统计总览')

headers7 = ["指标","数值"]
style_title_and_header(ws8, "数据统计总览", headers7)

#for ci, h in enumerate(['指标', '数值'], 1):
#    style_cell(ws8.cell(1, ci, h), h, None, 1, is_header=True)

s_ht = sum(r.get('合同金额', 0) for r in data if isinstance(r.get('合同金额'), (int, float)))
s_kp = sum(r.get('开票金额', 0) for r in data if isinstance(r.get('开票金额'), (int, float)))
s_hk = sum(r.get('回款合计', 0) for r in data if isinstance(r.get('回款合计'), (int, float)))
s_kpwr = sum(r.get('开票未回款', 0) for r in data if isinstance(r.get('开票未回款'), (int, float)))
rk = s_hk / s_kp if s_kp > 0 else 0
nc = len(set(r.get('合同号') for r in data))

items = [
    ('总合同金额', s_ht), ('总开票金额', s_kp), ('总回款合计', s_hk),
    ('总开票未回款', s_kpwr), ('整体回款率', rk),
    ('总合同数', nc), ('总数据笔数', len(data)), ('未回款原因分类数', len(rr))
]

ri = 3
for label, val in items:
    c1 = ws8.cell(ri, 1, label)
    c1.border = LIGHT_BORDER
    c1.alignment = CENTER_ALIGN
    c1.fill = ZEBRA_ODD if ri % 2 == 0 else ZEBRA_EVEN

    c2 = ws8.cell(ri, 2, val)
    c2.border = LIGHT_BORDER
    if label == '整体回款率':
        c2.number_format = '0.00%'
    elif label in ['总合同金额', '总开票金额', '总回款合计', '总开票未回款']:
        c2.number_format = '#,##0.000'
    c2.alignment = NUM_ALIGN
    c2.fill = ZEBRA_ODD if ri % 2 == 0 else ZEBRA_EVEN
    ri += 1

ws8.freeze_panes = 'A2'
ws8.column_dimensions['A'].width = 20
ws8.column_dimensions['B'].width = 20

# ============================================================
# 8. 保存
# ============================================================
wb_out.save(output_path)
print(f"Skill执行完成！文件已保存: {output_path}")
print(f"工作表: {wb_out.sheetnames}")
print(f"  原始数据: {len(data)} 行")
print(f"  同合同号合并汇总: {len(groups)} 行")
print(f"  未回款原因分类汇总: {len(rr)} 行")
print(f"  统计总览: {len(items)} 项")
