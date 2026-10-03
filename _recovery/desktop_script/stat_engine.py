# -*- coding: utf-8 -*-
"""
====通用明细统计主引擎（Skill内核）====
【重要约束说明！！！】
本文件禁止写入任何业务字段名称；禁止硬编码业务指标、维度、sheet工作表名称；
全部业务规则由外部JSON配置解析驱动；
统一调度流程：
  1、加载json配置；调用配置校验器做schema校验；
  2、读取原始Excel数据源；
  3、执行任务1：主键分组多行合并；
  4、执行任务2：专项分类汇总统计；
  5、执行任务3：多维度统计 + 全局总览指标计算；
  6、把所有结果DataFrame输出到多sheet的Excel；调用excel_styler执行统一美化；
更换业务场景，只需要新增一份业务JSON模板，内核文件无需改动。
"""
import json
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from agg_strategy import AGG_STRATEGY_REGISTRY, distinct_join
from config_validator import validate_business_config
from excel_styler import apply_sheet_format

class GeneralStatSkillEngine:
    def __init__(self, config_file_path: str):
        """
        引擎构造函数：加载并校验业务配置
        :param config_file_path: 业务模板json文件路径
        """
        # 读取json业务配置文件
        with open(config_file_path, encoding="utf-8") as f:
            self.config = json.load(f)
        # 执行配置schema合法性校验
        validate_business_config(self.config)
        # 提取输入配置与样式配置保存成员变量
        self.input_cfg = self.config["input"]
        self.style_cfg = self.config["style_setting"]
        # 原始数据集DataFrame对象，后续加载之后赋值
        self.df_raw = None
        self.input_excel_path = None

    def load_source_data(self, input_excel_path: str):
        """
        读取原始业务excel数据源；根据配置指定跳过前置标题行；
        :param input_excel_path:原始业务Excel文件路径
        """
        # header_row是业务数据表头所在行号；skiprows跳过前面N‑1行标题
        skip_rows = self.input_cfg["header_row"] - 1
        self.df_raw = pd.read_excel(
            input_excel_path,
            sheet_name=self.input_cfg["sheet_name"],
            skiprows=skip_rows
        )
        # 新增：列名清洗，去除列名前后空格，规避表头隐藏空格导致KeyError
        self.df_raw.columns = [str(col).strip() for col in self.df_raw.columns]
        self.input_excel_path = input_excel_path

    def run_task1_group_merge(self):
        """
        执行任务1：按照业务主键进行分组，多行明细合并为单行记录；
        聚合策略完全读取配置agg_strategy_list；支持distinct_join特殊拼接策略；
        并且插入配置指定的新连续序号列。
        :return: task1输出结果DataFrame
        """
        t1_cfg = self.config["task1_group_merge"]
        pk = t1_cfg["primary_key"]
        agg_list = t1_cfg["agg_strategy_list"]
        agg_def = dict()
        # 循环解析每一个字段对应的聚合策略
        for rule_item in agg_list:
            fld = rule_item["field_name"]
            strat = rule_item["agg_strategy"]
            # distinct_join属于带附加参数的特殊聚合策略，读取分隔符
            if strat == "distinct_join":
                sep = rule_item.get("join_separator", "；")
                agg_def[fld] = lambda s, sp=sep: distinct_join(s, sp)
            else:
                # 普通策略直接从注册表取函数
                agg_def[fld] = AGG_STRATEGY_REGISTRY[strat]
        # 根据业务主键分组并且执行聚合
        df_merged = self.df_raw.groupby(pk, as_index=False).agg(agg_def)
        df_merged.reset_index(drop=True, inplace=True)
        # 在第一列插入全新自增连续序号列，列名取自配置
        df_merged.insert(0, t1_cfg["new_serial_name"], range(1, len(df_merged)+1))
        return df_merged

    def run_task2_special_agg(self):
        """
        执行任务2：专项分类汇总；例如：未回款原因分类统计；
        指标定义、排序字段、占比衍生字段全部读取配置；执行降序排序。
        :return: task2结果DataFrame
        """
        t2_cfg = self.config["task2_special_agg"]
        grp_fld = t2_cfg["group_field"]
        amt_fld = t2_cfg["amount_field"]
        key_fld = t2_cfg["contract_key"]
        metrics_def = t2_cfg["metrics"]
        # 分组聚合计算基础指标（pd.NamedAgg兼容新版pandas）
        agg_list = []
        for out_col, cfg in metrics_def.items():
            src_col, agg_func = cfg
            agg_list.append(pd.NamedAgg(column=src_col, aggfunc=agg_func))
        df_grp = self.df_raw.groupby(grp_fld).agg(**{out_col: item for out_col,item in zip(metrics_def.keys(), agg_list)}).reset_index()
        # 根据配置中的目标汇总字段，计算整体占比（衍生指标）
        total_amt = df_grp[t2_cfg["target_sum_field"]].sum()
        df_grp[t2_cfg["ratio_field_name"]] = df_grp[t2_cfg["target_sum_field"]] / total_amt
        # 按照配置指定字段降序排序
        sort_col = t2_cfg["sort_by_field"]
        df_grp.sort_values(by=sort_col, ascending=False, inplace=True)
        df_grp.reset_index(drop=True, inplace=True)
        return df_grp

    def run_task3_multi_dim_and_overview(self):
        """
        执行任务3：多维度统计分析 + 全局总览指标；
        维度列表、基础指标、衍生比率指标、总览指标全部读取json配置；
        :return: dim_result_dict(各个维度sheet结果字典), df_overview(全局总览DataFrame)
        """
        t3_cfg = self.config["task3_multi_dim"]
        dim_result_dict = {}
        metric_map = t3_cfg["dim_metrics"]

        # 循环遍历每一个分析维度，执行分组统计
        for dim_item in t3_cfg["dim_list"]:
            dim_fld = dim_item["dimension_field"]
            sheet_n = dim_item["sheet_name"]
            metric_named = {}
            for out_col, setting in metric_map.items():
                src_col, agg_func = setting
                metric_named[out_col] = pd.NamedAgg(column=src_col, aggfunc=agg_func)
            df_dim = self.df_raw.groupby(dim_fld).agg(**metric_named).reset_index()

            # ==========【修改点1：维度派生比率，安全除法，全部读配置，无硬编码】==========
            for derive in t3_cfg["derived_ratio_list"]:
                out_col = derive["output_field"]
                numerator = derive["numerator_field"]
                denominator = derive["denominator_field"]

                def _safe_div(row):
                    num_val = row[numerator]
                    den_val = row[denominator]
                    if pd.isna(den_val) or float(den_val) == 0:
                        return np.nan
                    return num_val / den_val

                df_dim[out_col] = df_dim.apply(_safe_div, axis=1)
            # 使用配置指定的排序字段降序
            df_dim.sort_values(by=t3_cfg["sort_by_field"], ascending=False, inplace=True)
            dim_result_dict[sheet_n] = df_dim

        # --------【修改点2：全局统计总览，拆两轮循环，先普通聚合，后ratio比率计算】--------
        overview_dict = {}
        ov_list = t3_cfg["overview_metrics"]
        # 第一轮：普通聚合 sum / nunique / count_rows
        for ov in ov_list:
            out_n = ov["output_field"]
            agg_op = ov["agg_operator"]
            src_fld = ov.get("source_field", "")
            if agg_op == "sum":
                overview_dict[out_n] = self.df_raw[src_fld].sum()
            elif agg_op == "nunique":
                overview_dict[out_n] = self.df_raw[src_fld].nunique()
            elif agg_op == "count_rows":
                overview_dict[out_n] = len(self.df_raw)
            # ratio 本轮跳过，放到第二轮

        # 第二轮：专门处理ratio，从overview_dict读取已计算完成的指标，不再访问df_raw
        for ov in ov_list:
            agg_op = ov["agg_operator"]
            if agg_op == "ratio":
                out_n = ov["output_field"]
                num_key = ov["numerator"]
                den_key = ov["denominator"]
                num_val = overview_dict.get(num_key, np.nan)
                den_val = overview_dict.get(den_key, np.nan)
                if pd.isna(den_val) or float(den_val) == 0:
                    overview_dict[out_n] = np.nan
                else:
                    overview_dict[out_n] = num_val / den_val

        df_overview = pd.DataFrame([overview_dict])
        return dim_result_dict, df_overview

    def render_export_excel(self, output_file_path: str):
        """
        统一输出渲染excel；组装所有任务结果；循环创建工作表；调用样式美化；保存输出文件
        :param output_file_path:输出excel完整路径
        :return:输出文件路径字符串
        """
        out_sheet_cfg = self.config["output_sheet_names"]
            # 以原始工作簿为底稿，确保原始数据页的内容、顺序与单元格格式不变
            wb = load_workbook(self.input_excel_path)
            wb[self.input_cfg["sheet_name"]].title = out_sheet_cfg["raw_copy"]
        # 顺序执行全部统计任务，拿到结果数据集
        df_raw_copy = self.df_raw.copy()
        df_t1 = self.run_task1_group_merge()
        df_t2 = self.run_task2_special_agg()
        dim_result_map, df_overview = self.run_task3_multi_dim_and_overview()
        t3_cfg = self.config["task3_multi_dim"]
        # 先组装前3个基础sheet：原始数据、task1合并结果、task2专项汇总
        sheet_render_list = [
              (out_sheet_cfg["raw_copy"], df_raw_copy, [], [], [], []),
            (out_sheet_cfg["task1_result"], df_t1,
             self.config["task1_group_merge"]["num_fields"],
             self.config["task1_group_merge"]["pct_fields"],
             self.config["task1_group_merge"]["date_fields"],
             self.config["task1_group_merge"]["cat_fields"]),
            (self.config["task2_special_agg"]["sheet_name"], df_t2,
             self.config["task2_special_agg"]["num_fields"],
             self.config["task2_special_agg"]["pct_fields"],
             [],
             self.config["task2_special_agg"]["cat_fields"])
        ]
        # 追加循环生成各个维度分析sheet
        for sh_name, df_dim in dim_result_map.items():
            sheet_render_list.append((sh_name, df_dim,
                                     t3_cfg["num_fields"],
                                     t3_cfg["pct_fields"],
                                     [],
                                     t3_cfg["cat_fields"]))
        # 追加全局统计总览sheet
        sheet_render_list.append((t3_cfg["overview_sheet_name"], df_overview,
                                  t3_cfg["overview_num_fields"],
                                  t3_cfg["overview_pct_fields"], [], []))
        # 循环渲染每一张工作表，写入数据并应用样式美化
        for sheet_name, df_s, num_cols, pct_cols, date_cols, cat_cols in sheet_render_list:
            ws = wb.create_sheet(title=sheet_name)
            for row_data in dataframe_to_rows(df_s, index=False, header=True):
                ws.append(row_data)
            apply_sheet_format(ws, self.style_cfg, num_cols, pct_cols, date_cols, cat_cols)
        # 保存输出文件
        wb.save(output_file_path)
        return output_file_path