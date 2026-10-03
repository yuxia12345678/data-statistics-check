# -*- coding: utf-8 -*-
"""
Skill主业务引擎 stat_engine
核心设计：规则全部来自contract_repayment.json配置文件，业务逻辑不硬编码；
1、openpyxl复制原始工作表，完整保留原始内容、格式；
2、使用formula_builder生成全部Excel动态公式，输出公式而不是内存计算静态值；
3、任务1：合同号分组合并；任务2：未回款原因分类汇总；任务3：4个维度+统计总览；
4、调用excel_styler完成全部工作表统一美化；
5、可扩展：更换JSON配置，即可适配其他业务，无需修改内核代码。
"""
import json
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from agg_strategy import AGG_STRATEGY_REGISTRY
from config_validator import validate_business_config
from excel_styler import apply_sheet_format
from formula_builder import (
    build_sumifs_formula, build_countifs_formula,
    build_sumproud_distinct_contract_formula, build_divide_formula,
    build_textjoin_distinct_formula, build_index_first_non_null
)


class GeneralStatSkillEngine:
    def __init__(self, config_file_path: str):
        """初始化引擎：加载JSON配置并做schema校验"""
        with open(config_file_path, encoding="utf-8") as f:
            self.config = json.load(f)
        validate_business_config(self.config)
        self.input_cfg = self.config["input"]
        self.style_cfg = self.config["style_setting"]
        self.output_sheet_cfg = self.config["output_sheet_names"]
        self.sheet_title_cfg = self.config["sheet_titles"]
        self.t1_cfg = self.config["task1_group_merge"]
        self.t2_cfg = self.config["task2_special_agg"]
        self.t3_cfg = self.config["task3_multi_dim"]
        self.df_raw = None
        self.input_excel_path = ""

    # def load_source_data(self, input_excel_path: str):
    #     """
    #     读取输入Excel；pandas仅用于提取维度唯一值，用于生成公式行；
    #     真正输出原始工作表：使用openpyxl直接复制源文件，保证原始格式不丢失。
    #     """
    #     self.input_excel_path = input_excel_path
    #     skip_rows = self.input_cfg["header_row"] - 1
    #     self.df_raw = pd.read_excel(
    #         input_excel_path,
    #         sheet_name=self.input_cfg["sheet_name"],
    #         skiprows=skip_rows
    #     )
    #     self.df_raw.columns = [str(col).strip() for col in self.df_raw.columns]

    def load_source_data(self, input_excel_path: str):
        """
        读取输入Excel数据源。
        除了基础的读取操作外，这里还重点加入了数据清洗逻辑，
        以解决由于“合同号”存在隐藏空格或换行符导致的统计偏差。
        """
        # 保存输入文件路径到实例属性，供后续渲染及导出时使用
        self.input_excel_path = input_excel_path       
        # 因为 Python/Pandas 索引是从0开始的，而 Excel 行号从1开始，所以要减 1
        skip_rows = self.input_cfg["header_row"] - 1      
        # 使用 Pandas 读取 Excel 文件
        self.df_raw = pd.read_excel(
            input_excel_path,
            sheet_name=self.input_cfg["sheet_name"], # 指定读取的 Sheet 名
            skiprows=skip_rows                        # 跳过表头之前的无用行（如大标题、“单位”行）
        )       
        # 清理列名：去除列名首尾的空格，并统一转换为字符串类型
        self.df_raw.columns = [str(col).strip() for col in self.df_raw.columns]     
        # --- 核心修复：强制清理“合同号”列的空格和不可见字符 ---
        contract_pk = self.t1_cfg["primary_key"]   
        # 确保该字段确实存在于数据表中，防止抛错
        if contract_pk in self.df_raw.columns:
            # 逐步处理“合同号”列：
            # 1. .astype(str)      强制转换为字符串，防止数字或日期被识别为其他类型
            # 2. .str.strip()      去掉首尾的空格（如 " HT20250101 " -> "HT20250101"）
            # 3. .str.replace(...) 去掉内部可能存在的换行符 \\n 和回车符 \\r
            # 注意：这里单引号内只需要写 \\n 和 \\r，不要多加斜杠写成 \\\\n，否则无法匹配到真正的换行符
            self.df_raw[contract_pk] = (
                self.df_raw[contract_pk]
                .astype(str)
                .str.strip()
                .str.replace('\\n', '')
                .str.replace('\\r', '')
            )
        # ------------------------------------------------

    @staticmethod
    def build_header_col_mapping(raw_ws, header_row: int):
        """读取原始工作表表头行，构建 {字段名:Excel大写列字母}"""
        mapping = {}
        for col_idx, cell in enumerate(raw_ws[header_row], start=1):
            field_name = str(cell.value).strip() if cell.value is not None else ""
            mapping[field_name] = get_column_letter(col_idx)
        return mapping

    def get_unique_dim_list(self, dim_field: str):
        """获取维度去重列表，用于生成公式行，仅内存读取，不参与最终输出值计算"""
        return sorted(self.df_raw[dim_field].dropna().unique().tolist())

    def get_sorted_dim_list(self, dim_field: str, sort_field: str = None):
        """获取维度去重列表，如果配置了排序字段，则按该字段汇总降序排列"""
        if sort_field and sort_field in self.df_raw.columns:
            # 按维度分组，并对排序字段求和
            grouped = self.df_raw.groupby(dim_field)[sort_field].sum().reset_index()
            # 按求和结果降序排列
            grouped = grouped.sort_values(by=sort_field, ascending=False)
            return grouped[dim_field].tolist()
        else:
            return self.get_unique_dim_list(dim_field)

    def run_task1_formula_build(self, wb: Workbook, field_to_col: dict):
        """
        任务1：同合同号合并汇总工作表，全部写入Excel动态公式
        :param wb: 输出workbook对象
        :param field_to_col: dict {字段名:列字母} 从原始表头解析
        """
        sheet_name = self.output_sheet_cfg["task1_result"]
        ws = wb.create_sheet(title=sheet_name)
        # 表头全部取自配置，序号名称读取json
        header_list = [self.t1_cfg["new_serial_name"]] + [item["field_name"] for item in self.t1_cfg["agg_strategy_list"]]
        # 表头写入第2行，第1行留给合并大标题
        ws.append([""])
        ws.append(header_list)

        contract_pk = self.t1_cfg["primary_key"]
        contract_list = self.get_unique_dim_list(contract_pk)
        pk_col_letter = field_to_col[contract_pk]

        # 遍历每一个合同号，逐行写入公式
        for serial_no, contract_val in enumerate(contract_list, start=1):
            row_data = [serial_no]
            for rule_item in self.t1_cfg["agg_strategy_list"]:
                field_name = rule_item["field_name"]
                strat = rule_item["agg_strategy"]
                col_letter = field_to_col.get(field_name)
                if not col_letter:
                    raise ValueError(f"【Task1】原始数据表头找不到字段：{field_name}")

                if strat == "sum":
                    formula = build_sumifs_formula(
                        sum_col=col_letter,
                        cond_map={pk_col_letter: f'"{contract_val}"'}
                    )
                    row_data.append(formula)
                elif strat == "first_non_null":
                    formula = build_index_first_non_null(
                        group_col=pk_col_letter, group_val_cell=f'"{contract_val}"', fetch_col=col_letter
                    )
                    row_data.append(formula)
                elif strat == "distinct_join":
                    sep = rule_item.get("join_separator", "；")
                    formula = build_textjoin_distinct_formula(
                        group_col=pk_col_letter, group_val_cell=f'"{contract_val}"',
                        target_col=col_letter, sep=sep
                    )
                    row_data.append(formula)
                elif strat == "value":
                    row_data.append(contract_val)
                else:
                    row_data.append("")
            ws.append(row_data)

        apply_sheet_format(
            ws, self.style_cfg,
            num_cols=self.t1_cfg["num_fields"],
            pct_cols=self.t1_cfg["pct_fields"],
            date_cols=self.t1_cfg["date_fields"],
            cat_cols=self.t1_cfg["cat_fields"],
            header_row_idx=2,
            sheet_title=self.sheet_title_cfg["task1_result"],
            total_title_cols=len(header_list)
        )
        return ws

    def run_task2_formula_build(self, wb: Workbook, field_to_col: dict):
        """
        任务2：未回款原因分类汇总工作表，全部Excel动态公式；开启分类条件格式
        表头、指标名称全部读取json配置，消除硬编码字符串
        """
        sheet_name = self.t2_cfg["sheet_name"]
        ws = wb.create_sheet(title=sheet_name)
        # 表头全部读取配置，不再写死字面量；第1行留给大标题
        header_list = [
            "序号",
            self.t2_cfg["group_field"],
            *list(self.t2_cfg["metrics"].keys()),
            self.t2_cfg["ratio_field_name"]
        ]
        ws.append([""])
        ws.append(header_list)

        group_filed = self.t2_cfg["group_field"]
        amount_field = self.t2_cfg["amount_field"]
        contract_key = self.t2_cfg["contract_key"]

        group_col_letter = field_to_col[group_filed]
        amt_sum_col = field_to_col[amount_field]
        contract_col = field_to_col[contract_key]

        group_val_list = self.get_unique_dim_list(group_filed)
        for serial_no, g_val in enumerate(group_val_list, start=1):
            cell_ref = f'"{g_val}"'
            sum_formula = build_sumifs_formula(amt_sum_col, cond_map={group_col_letter: cell_ref})
            contract_cnt_formula = build_sumproud_distinct_contract_formula(
                group_col=group_col_letter, val_cell=cell_ref, contract_col=contract_col
            )
            row_cnt_formula = build_countifs_formula(cond_map={group_col_letter: cell_ref})
            ws.append([serial_no, g_val, sum_formula, contract_cnt_formula, row_cnt_formula, ""])

        # 占比列，通过表头动态获取占比所在列，不再写死F列
        header_row_cells = [c.value for c in ws[2]]
        ratio_col_index = header_row_cells.index(self.t2_cfg["ratio_field_name"]) + 1
        ratio_col_letter = get_column_letter(ratio_col_index)
        value_col_index = header_row_cells.index(list(self.t2_cfg["metrics"].keys())[0]) + 1
        value_col_letter = get_column_letter(value_col_index)

        max_row = ws.max_row
        for r in range(3, max_row + 1):
            ws[f"{ratio_col_letter}{r}"].value = f'=IFERROR({value_col_letter}{r}/SUM(${value_col_letter}$3:${value_col_letter}${max_row}),"")'

        apply_sheet_format(
            ws, self.style_cfg,
            num_cols=self.t2_cfg["num_fields"],
            pct_cols=self.t2_cfg["pct_fields"],
            int_cols=self.t2_cfg.get("int_fields", []),  # <--- 新增整数传入
            date_cols=[],
            cat_cols=[self.t2_cfg["group_field"]],
            header_row_idx=2,
            enable_cat_cond_format=True,
            cat_field_name=self.t2_cfg["group_field"],
            sheet_title=self.sheet_title_cfg["task2"],
            total_title_cols=len(header_list)
        )
        return ws

    #def run_task3_multi_dim_formula_build(self, wb: Workbook, field_to_col: dict):
    def run_task3_multi_dim_formula_build(self, wb: Workbook, field_to_col: dict, field_to_col_t1: dict):
        """
        任务3：4个维度统计工作表 + 统计总览工作表，全部输出Excel动态公式
        """
        t3_cfg = self.t3_cfg
        dim_result_sheet_map = {}
        derived_ratio_list = self.t3_cfg["derived_ratio_list"]

        for dim_item in t3_cfg["dim_list"]:
            sheet_n = dim_item["sheet_name"]
            dim_field = dim_item["dimension_field"]
            dim_col_letter = field_to_col[dim_field]
            ws = wb.create_sheet(title=sheet_n)
            header = ["序号", dim_field] + list(t3_cfg["dim_metrics"].keys()) + [x["output_field"] for x in derived_ratio_list]
            ws.append([""])
            ws.append(header)
            #dim_val_list = self.get_unique_dim_list(dim_field)

            # 读取配置中的排序字段，并调用新的排序方法
            sort_field = t3_cfg.get("sort_by_field", None)
            dim_val_list = self.get_sorted_dim_list(dim_field, sort_field)

            for serial_no, d_val in enumerate(dim_val_list, start=1):
                cond_cell = f'"{d_val}"'
                row_buf = [serial_no, d_val]
                for _, setting in t3_cfg["dim_metrics"].items():
                    src_field, agg_type = setting
                    src_col = field_to_col[src_field]
                    if agg_type == "sum":
                        f = build_sumifs_formula(src_col, cond_map={dim_col_letter: cond_cell})
                        row_buf.append(f)
                    # elif agg_type == "nunique":
                    #     # 增加对“去重计数（合同数）”的处理
                    #     f = build_sumproud_distinct_contract_formula(
                    #         group_col=dim_col_letter, 
                    #         val_cell=cond_cell, 
                    #         contract_col=src_col
                    #     )
                    #     row_buf.append(f)
                    elif agg_type == "nunique":
                        # 直接引用汇总表的列（例如：同合同号合并汇总!$B:$B）
                        t1_sheet = self.output_sheet_cfg["task1_result"]
                        # 动态获取该字段在汇总表中的列字母
                        src_col_t1 = field_to_col_t1[src_field]    
                        # 采用通用 COUNTIFS 公式，完美处理空值，速度极快
                        # f = f'=COUNTIFS(\\'{t1_sheet}\\'!${src_col_t1}:${src_col_t1},{cond_cell},\\'{t1_sheet}\\'!${src_col_t1}:${src_col_t1},"<>")'
                        f = f"=COUNTIFS('{t1_sheet}'!${src_col_t1}:${src_col_t1},{cond_cell},'{t1_sheet}'!${src_col_t1}:${src_col_t1},\"<>\")"
                        row_buf.append(f)
                    else:
                        row_buf.append("")
                # 比率单元格先占位，后面填充公式
                for _item in derived_ratio_list:
                    row_buf.append("")
                ws.append(row_buf)

            data_max_r = ws.max_row
            header_cells = [c.value for c in ws[2]]
            # # 循环派生比率配置，动态定位列，不再写死H/I
            # for ratio_cfg in derived_ratio_list:
            #     out_name = ratio_cfg["output_field"]
            #     num_field = ratio_cfg["numerator_field"]
            #     den_field = ratio_cfg["denominator_field"]
            #     col_idx = header_cells.index(out_name) + 1
            #     col_letter = get_column_letter(col_idx)
            #     num_col_idx = header_cells.index(num_field) + 1
            #     den_col_idx = header_cells.index(den_field) + 1
            #     num_let = get_column_letter(num_col_idx)
            #     den_let = get_column_letter(den_col_idx)
            #     for r in range(3, data_max_r + 1):
            #         ws[f"{col_letter}{r}"].value = build_divide_formula(f"{num_let}{r}", f"{den_let}{r}")
            # 循环派生比率配置，动态定位列，不再写死H/I
            for ratio_cfg in derived_ratio_list:
                out_name = ratio_cfg["output_field"]
                num_field = ratio_cfg["numerator_field"]
                den_field = ratio_cfg["denominator_field"]
                col_idx = header_cells.index(out_name) + 1
                col_letter = get_column_letter(col_idx)
                num_col_idx = header_cells.index(num_field) + 1
                num_let = get_column_letter(num_col_idx)
                # 判断分母是否为“全局总合同金额”
                if den_field == "全局总合同金额":
                    # 从 field_to_col 动态获取“合同金额”所在的列字母（通常为F）
                    contract_amt_col = field_to_col.get("合同金额", "F")
                    # 构建全局求和公式字符串，例如：SUM(原始数据!$F:$F)
                    den_expr = f"SUM(原始数据!${contract_amt_col}:${contract_amt_col})"
                    for r in range(3, data_max_r + 1):
                        ws[f"{col_letter}{r}"].value = build_divide_formula(f"{num_let}{r}", den_expr)
                else:
                    # 正常的本表内分母（如回款率）
                    den_col_idx = header_cells.index(den_field) + 1
                    den_let = get_column_letter(den_col_idx)
                    for r in range(3, data_max_r + 1):
                        ws[f"{col_letter}{r}"].value = build_divide_formula(f"{num_let}{r}", f"{den_let}{r}")

            apply_sheet_format(
                ws, self.style_cfg,
                num_cols=t3_cfg["num_fields"],
                pct_cols=t3_cfg["pct_fields"],
                int_cols=t3_cfg.get("int_fields", []),  # <--- 新增整数传入
                date_cols=[],
                cat_cols=[dim_field],
                header_row_idx=2,
                sheet_title=f"{dim_field}{self.sheet_title_cfg['dim_suffix']}",
                total_title_cols=len(header)
            )
            dim_result_sheet_map[sheet_n] = ws

        # ==========生成统计总览工作表【补全①：完整动态公式】==========
        overview_sheet_name = t3_cfg["overview_sheet_name"]
        ws_over = wb.create_sheet(title=overview_sheet_name)
        over_header = ["指标", "数值"]
        ws.append([""])
        ws_over.append([""])
        ws_over.append(over_header)
        over_metrics = t3_cfg["overview_metrics"]
        cell_of = dict()
        start_data_row = 3
        # --- 新增：动态获取原始数据表的名称和有效行范围 ---
        raw_sheet_name = self.output_sheet_cfg["raw_copy"]  # 获取配置中的"原始数据"表名
        start_row = 4                                       # 数据从第4行开始（大标题占1行，表头占2行）
        end_row = wb[raw_sheet_name].max_row                # 动态获取该表的最大有效行数
        # ------------------------------------------------
        for idx, item in enumerate(over_metrics):
            cur_row = start_data_row + idx
            output_name = item["output_field"]
            agg_op = item["agg_operator"]
            src_field = item.get("source_field", "")
            divide_coeff = item.get("divide", 1)
            numerator_name = item.get("numerator", "")
            denominator_name = item.get("denominator", "")
            formula = ""
            src_col = field_to_col.get(src_field, "")

            if agg_op == "sum":
                if divide_coeff and divide_coeff != 1:
                    formula = f'=SUM(原始数据!${src_col}:${src_col})/{divide_coeff}'
                else:
                    formula = f'=SUM(原始数据!${src_col}:${src_col})'
            # elif agg_op == "nunique":
            #     # SUMPRODUCT不重复计数
            #     formula = f'=SUMPRODUCT(1/COUNTIF(原始数据!${src_col}:${src_col},原始数据!${src_col}:${src_col}))'
            # elif agg_op == "nunique":
            #     # 修复：加入容错，防止空白单元格导致 #DIV/0!
            #     countif_expr = f'COUNTIF(原始数据!${src_col}:${src_col},原始数据!${src_col}:${src_col})'
            #     formula = f'=SUMPRODUCT((原始数据!${src_col}:${src_col}<>"")/IF({countif_expr}=0,1,{countif_expr}))'
            # elif agg_op == "count_rows":
            #     formula = f'=COUNTA(原始数据!A:A)-1'
            elif agg_op == "nunique":
                # 动态引用任务1汇总表，减1是为了去掉表头
                t1_sheet = self.output_sheet_cfg["task1_result"]
                src_col_t1 = field_to_col_t1.get(src_field, "B")
                #formula = f'=COUNTA(\\\\'{t1_sheet}\\\\'!${src_col_t1}:${src_col_t1})-1'
                formula = f"=COUNTA('{t1_sheet}'!${src_col_t1}:${src_col_t1})-1"
            elif agg_op == "count_rows":
                # 严格限定数据区域，避免包含大标题和表头
                formula = f"=COUNTA('{raw_sheet_name}'!A{start_row}:A{end_row})"
            elif agg_op == "ratio":
                # 从cell_of字典拿到前面已经生成的行号
                num_r = cell_of[numerator_name]
                den_r = cell_of[denominator_name]
                formula = f'=IFERROR(B{num_r}/B{den_r},"")'
            ws_over.append([output_name, formula])
            cell_of[output_name] = cur_row
            # 针对总览表纵向布局，动态应用数字格式
            val_cell = ws_over.cell(row=cur_row, column=2)
            if output_name in t3_cfg["overview_num_fields"]:
                val_cell.number_format = self.style_cfg["overview_number_format"]
            elif output_name in t3_cfg["overview_pct_fields"]:
                val_cell.number_format = self.style_cfg["percent_number_format"]

        apply_sheet_format(
            ws_over, self.style_cfg,
            num_cols=[],
            pct_cols=[],
            date_cols=[],
            cat_cols=["指标"],
            header_row_idx=2,
            sheet_title=self.sheet_title_cfg["overview"],
            total_title_cols=len(over_header)
        )
        return dim_result_sheet_map, ws_over

    def render_export_excel(self, output_file_path: str):
        """
        统一输出入口
        1、openpyxl加载源Excel，复用源workbook内存对象；修改sheet名称；
        2、处理原始数据表的大标题
        3、依次执行任务1、任务2、任务3，全部写入动态公式；
        4、保存输出文件。
        """
        # src_wb = load_workbook(self.input_excel_path, data_only=False)
        # src_raw_ws = src_wb[self.input_cfg["sheet_name"]]
        # # 修改工作表名称
        # src_raw_ws.title = self.output_sheet_cfg["raw_copy"]
        # header_row = self.input_cfg["header_row"]
        # field_to_col_mapping = self.build_header_col_mapping(src_raw_ws, header_row)

        # # 给原始数据表设置A1合并大标题，原始表表头不变，插入空行作为第1行
        # src_raw_ws.insert_rows(1)
        # raw_header_cols = len([c.value for c in src_raw_ws[2]])
        # apply_sheet_format(
        #     src_raw_ws, self.style_cfg,
        #     num_cols=self.t1_cfg["num_fields"],
        #     pct_cols=[],
        #     date_cols=self.t1_cfg["date_fields"],
        #     cat_cols=self.t1_cfg["cat_fields"],
        #     header_row_idx=2,
        #     sheet_title=self.sheet_title_cfg["raw_copy"],
        #     total_title_cols=raw_header_cols
        # )

        # 给原始数据表设置A1合并大标题，原始表表头不变，插入空行作为第1行
        #src_raw_ws.insert_rows(1)

        src_wb = load_workbook(self.input_excel_path, data_only=False)
        src_raw_ws = src_wb[self.input_cfg["sheet_name"]]

        # 修改工作表名称
        src_raw_ws.title = self.output_sheet_cfg["raw_copy"]

        # # 删除原始数据表中的第2行（单位：万元）
        # # 执行删除后，原第3行的表头会变成第2行
        # if src_raw_ws.max_row >= 2:
        #     src_raw_ws.delete_rows(2)
        
        # 先判断第二行第一列的值是否包含“单位”，包含则删除原始数据表中的第2行（单位：万元）
        cell_val = src_raw_ws.cell(row=2, column=1).value
        if cell_val and "单位" in str(cell_val):
            src_raw_ws.delete_rows(2)  # 如果是“单位”行才删
            
        header_row = self.input_cfg["header_row"]  # 此时 header_row 仍为 2（指向真正的表头）
        field_to_col_mapping = self.build_header_col_mapping(src_raw_ws, header_row)

        # 给原始数据表插入空行作为第1行的大标题
        src_raw_ws.insert_rows(1)

        # 动态计算新的表头行索引（原表头行号 + 1）
        new_header_row_idx = header_row + 1
        #raw_header_cols = len([c.value for c in src_raw_ws[new_header_row_idx]])

        # 只统计连续的非空表头数量，遇到空列立即停止
        raw_header_cols = 0
        for c in src_raw_ws[new_header_row_idx]:
            if c.value is not None and str(c.value).strip() != "":
                raw_header_cols += 1
            else:
                break
        apply_sheet_format(
            src_raw_ws, self.style_cfg,
            num_cols=self.t1_cfg["num_fields"],
            pct_cols=[],
            date_cols=self.t1_cfg["date_fields"],
            cat_cols=self.t1_cfg["cat_fields"],
            header_row_idx=new_header_row_idx,  # <--- 使用动态计算出的行号
            sheet_title=self.sheet_title_cfg["raw_copy"],
            total_title_cols=raw_header_cols
        )

        # # 依次运行三大业务任务，传入映射字典
        # self.run_task1_formula_build(src_wb, field_to_col_mapping)
        # self.run_task2_formula_build(src_wb, field_to_col_mapping)
        # self.run_task3_multi_dim_formula_build(src_wb, field_to_col_mapping)

        # 依次运行三大业务任务，传入映射字典
        # 1. 运行任务1，并拿到生成的 worksheet 对象
        ws_task1 = self.run_task1_formula_build(src_wb, field_to_col_mapping)
        # 2. 构建任务1（汇总表）的表头列映射
        # 任务1的表头在第2行（因为第1行是大标题）
        header_row_t1 = 2
        field_to_col_mapping_t1 = self.build_header_col_mapping(ws_task1, header_row_t1)
        # 3. 运行任务2（不需要汇总表映射）
        self.run_task2_formula_build(src_wb, field_to_col_mapping)
        # 4. 运行任务3（需要同时传入原始表映射和汇总表映射）
        self.run_task3_multi_dim_formula_build(src_wb, field_to_col_mapping, field_to_col_mapping_t1)

        src_wb.save(output_file_path)
        src_wb.close()
        return output_file_path