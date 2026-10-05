# -*- coding: utf-8 -*-
"""
Skill 主业务引擎 stat_engine（配置驱动版）

核心设计：业务规则全部来自 JSON 配置，代码不硬编码业务字段。
通用聚合走 agg_strategy.AGG_STRATEGY_REGISTRY，业务字段名全部来自配置。

本次修正（对齐题目规范）：
1. **「原始数据」零改动**：不再插入大标题行、不再删除「单位」行；值 / 行列顺序 /
   数字格式逐格保持源文件原样，仅叠加表头样式、斑马纹、水平边框、对齐、列宽、冻结；
2. **表顺序**：统计总览排最后
   （原始数据 / 同合同号合并汇总 / 未回款原因分类汇总 / 4 个维度表 / 统计总览）；
3. **序号公式化**：所有序号列写入 `=ROW()-2`，不再是字面量整数；
4. **合并表排序**：按合同号升序（从小到大）；
5. **合同数口径**：组内**不同合同号**个数（SUMPRODUCT 去重计数）；修正原先
   拿维度值去比对合同号列、导致合同数恒为 0 的错误；
6. **未回款原因分类数**：按非空去重计数；修正原先 COUNTA 把空文本公式计入的错误；
7. **公式区间限定**：统一传入数据行区间，不再整列引用；
8. **缓存值注入**：pandas 独立算出期望值，保存后写入 OOXML `<v>`，
   使产物既能被 Excel 重算，也能被程序读取。

两条路径互不复用：公式由 formula_builder 生成，期望值由本模块用 pandas 独立计算。
"""
import json
import os
import time

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from agg_strategy import (
    AGG_STRATEGY_REGISTRY, distinct_join, first_non_null, is_blank, text_of,
)
from cached_values import inject_cached_values, scan_errors
from config_validator import validate_business_config
from excel_styler import apply_sheet_format
from formula_builder import (
    build_sumifs_formula, build_countifs_formula,
    build_sumproud_distinct_contract_formula, build_divide_formula,
    build_divide_by_global_formula, build_divide_by_columns_diff_formula,
    build_index_first_non_null,
    build_seq_formula, build_sum_all_formula, build_counta_formula,
    build_distinct_all_formula, build_distinct_nonblank_all_formula,
)

AGG_SUM = AGG_STRATEGY_REGISTRY["sum"]
AGG_NUNIQUE = AGG_STRATEGY_REGISTRY["nunique"]

# 空白判定（is_blank）与文本规范化（text_of）已下沉到 agg_strategy 统一实现——
# 公式缓存值、分组键、排序键共用同一语义，避免两套判空逻辑漂移。


class GeneralStatSkillEngine:
    def __init__(self, config_file_path: str):
        """初始化引擎：加载 JSON 配置并做 schema 校验。"""
        with open(config_file_path, encoding="utf-8") as f:
            self.config = json.load(f)
        validate_business_config(self.config)
        self.config_path = config_file_path

        self.input_cfg = self.config["input"]
        self.style_cfg = self.config["style_setting"]
        self.output_sheet_cfg = self.config["output_sheet_names"]
        self.sheet_title_cfg = self.config["sheet_titles"]
        self.t1_cfg = self.config["task1_group_merge"]
        self.t2_cfg = self.config["task2_special_agg"]
        self.t3_cfg = self.config["task3_multi_dim"]

        self.df_raw = None
        self.input_excel_path = ""
        self.field_to_col = {}
        self.first_row = self.input_cfg["header_row"] + 1     # 数据首行（=3）
        self.last_row = None                                  # 数据末行，读取源表后确定
        self.expected = {}                                    # {sheet: {coord: 值}}
        self.report = {}                                      # 运行回执（由 run_skill.py 落盘）

    # ============================================================ 读取源数据

    def load_source_data(self, input_excel_path: str):
        """
        读取输入 Excel。
        pandas 仅用于提取维度唯一值、排序键与"期望值"，从而生成公式行并注入缓存值；
        真正输出的「原始数据」表由 openpyxl 直接沿用源工作表，保证原始格式不丢失。
        """
        self.input_excel_path = input_excel_path
        header_row = self.input_cfg["header_row"]

        wb = load_workbook(input_excel_path, data_only=True)
        ws = wb[self.input_cfg["sheet_name"]]
        self.field_to_col = self.build_header_col_mapping(ws, header_row)

        # 动态探测数据末行（最后一行有任意非空单元格）。
        # 用 iter_rows 单遍扫描代替逐格 ws.cell(r, c)：后者每次调用都要做一次
        # 坐标解析，在千行级源表上是读入阶段的主要开销；源文件自带的
        # 「幽灵列」（有样式无值）不影响判定（value 均为 None）。
        last = header_row
        for row in ws.iter_rows(min_row=header_row + 1):
            if any(cell.value not in (None, "") for cell in row):
                last = row[0].row
        self.last_row = last
        # 运行回执：源表侧的客观事实
        self.report.update({
            "input": os.path.abspath(input_excel_path),
            "config": os.path.abspath(self.config_path),
            "profile": "config-driven",
            "source_sheet": self.input_cfg["sheet_name"],
            "hidden_sheets": [w.title for w in wb.worksheets if w.sheet_state != "visible"],
            "header_row": header_row,
            "data_first_row": self.first_row,
            "data_last_row": self.last_row,
            "source_columns": len(self.field_to_col),
        })
        wb.close()

        df = pd.read_excel(input_excel_path,
                           sheet_name=self.input_cfg["sheet_name"],
                           skiprows=header_row - 1)
        df.columns = [str(c).strip() for c in df.columns]
        keep = [c for c in df.columns if c in self.field_to_col]
        self.df_raw = df[keep].copy()
        self.report["records"] = int(self.df_raw.notna().any(axis=1).sum())

        # 合同号清理：去首尾空白与内部换行，避免隐藏字符导致统计偏差
        pk = self.t1_cfg["primary_key"]
        if pk in self.df_raw.columns:
            self.df_raw[pk] = (self.df_raw[pk].astype(str)
                               .str.strip()
                               .str.replace("\n", "", regex=False)
                               .str.replace("\r", "", regex=False))

    @staticmethod
    def build_header_col_mapping(raw_ws, header_row: int) -> dict:
        """读取表头行，构建 {字段名: Excel 大写列字母}；空表头跳过，重复取首次出现。"""
        mapping = {}
        for col_idx in range(1, raw_ws.max_column + 1):
            v = raw_ws.cell(header_row, col_idx).value
            name = str(v).strip() if v is not None else ""
            if name and name not in mapping:
                mapping[name] = get_column_letter(col_idx)
        return mapping

    # ============================================================ 分组口径

    def _col(self, field: str) -> str:
        col = self.field_to_col.get(field)
        if not col:
            raise ValueError("原始数据表头找不到字段：%s" % field)
        return col

    def get_sorted_dim_list(self, dim_field: str, sort_field: str) -> list:
        """维度去重列表：按 sort_field 汇总降序（并列时保持首次出现顺序）。"""
        if sort_field and sort_field in self.df_raw.columns:
            grouped = self.df_raw.groupby(dim_field)[sort_field].sum()
            grouped = grouped.sort_values(ascending=False, kind="stable")
            return [k for k in grouped.index.tolist() if not is_blank(k)]
        vals = self.df_raw[dim_field].dropna().unique().tolist()
        return sorted([v for v in vals if not is_blank(v)])

    def _resolve_global_base_field(self, den_field: str) -> str:
        """把「全局<指标名>」形式的**跨维度分母**解析回它的**源字段**（零硬编码）。

        例：`den_field="全局总合同金额"` → 去掉「全局」→ 在 `dim_metrics` 里按输出名
        反查 `source_field` → `"合同金额"`。换一套业务字段（如「订单金额」）时只改配置，
        代码无需改动。

        :param den_field: 形如「全局总合同金额」的派生比率分母字段名
        :return: 对应的源字段名；无法解析时回落到任务2 的占比分母被减数
        """
        key = den_field[2:] if den_field.startswith("全局") else den_field
        setting = (self.t3_cfg.get("dim_metrics") or {}).get(key)
        if setting:
            return setting[0]
        den_cfg = self.t2_cfg.get("ratio_denominator") or {}
        return den_cfg.get("minuend_field", "")

    def get_contract_list(self) -> list:
        """合同号列表：按合同号升序（从小到大；按文本比较，结果稳定可复现）。"""
        pk = self.t1_cfg["primary_key"]
        vals = self.df_raw[pk].dropna().unique().tolist()
        return sorted([v for v in vals if not is_blank(v)], key=text_of)

    # ============================================================ 期望值记录

    def _put(self, ws, row: int, col: int, formula, expected=None):
        """写入单元格，并登记该公式的期望值（供缓存值注入）。"""
        cell = ws.cell(row=row, column=col)
        cell.value = formula
        if expected is not None:
            self.expected.setdefault(ws.title, {})[cell.coordinate] = expected

    # ============================================================ 任务1

    def run_task1_formula_build(self, wb, field_to_col: dict):
        """任务1：同合同号合并汇总。数值求和、文本取首条非空、备注去重拼接，全部动态公式。"""
        ws = wb.create_sheet(title=self.output_sheet_cfg["task1_result"])
        rules = self.t1_cfg["agg_strategy_list"]
        header_list = [self.t1_cfg["new_serial_name"]] + [r["field_name"] for r in rules]
        ws.append([""])
        ws.append(header_list)

        pk = self.t1_cfg["primary_key"]
        pk_col = self._col(pk)
        grouped = self.df_raw.groupby(pk, sort=False)
        contract_list = self.get_contract_list()
        self.report["contract_count"] = len(contract_list)

        # 循环外预计算每条规则的（字段名, 策略, Excel 列字母, 拼接分隔符）：
        # 下面是「合同数 × 规则数」的双重循环，配置查找与表头定位都提到循环外只做一次
        rule_plans = [
            (rule["field_name"], rule["agg_strategy"],
             self._col(rule["field_name"]), rule.get("join_separator", "；"))
            for rule in rules
        ]

        for i, code in enumerate(contract_list):
            row = self.first_row + i
            g = grouped.get_group(code)
            ref = '"%s"' % code
            self._put(ws, row, 1, build_seq_formula(self.first_row - 1), i + 1)
            for j, (field, strat, col, join_sep) in enumerate(rule_plans, start=2):
                if strat == "sum":
                    self._put(ws, row, j,
                              build_sumifs_formula(col, {pk_col: ref},
                                                   first_row=self.first_row, last_row=self.last_row),
                              float(AGG_SUM(g[field])))
                elif strat == "first_non_null":
                    exp = first_non_null(g[field])
                    self._put(ws, row, j,
                              build_index_first_non_null(pk_col, ref, col,
                                                         first_row=self.first_row,
                                                         last_row=self.last_row),
                              exp if not is_blank(exp) else "")
                elif strat == "distinct_join":
                    text = distinct_join(g[field], join_sep)
                    # 去重拼接是文本聚合：Excel 无稳定的非数组写法
                    # （TEXTJOIN/UNIQUE/FILTER 需 Excel365），故直接写入结果文本；
                    # verify.py 的 R3a 仅对「数值单元格必须是公式」做断言，
                    # 文本字面量不在其列（公式化例外）。
                    self._put(ws, row, j, text, text)
                elif strat == "value":
                    self._put(ws, row, j, code, None)
                else:
                    raise ValueError("【Task1】不支持的聚合策略：%s" % strat)

        apply_sheet_format(
            ws, self.style_cfg,
            num_cols=self.t1_cfg["num_fields"],
            pct_cols=self.t1_cfg["pct_fields"],
            date_cols=self.t1_cfg["date_fields"],
            cat_cols=self.t1_cfg["cat_fields"],
            int_cols=self.t1_cfg.get("int_fields", []),
            text_cols=self.t1_cfg.get("text_fields", []),
            header_row_idx=2,
            sheet_title=self.sheet_title_cfg["task1_result"],
            total_title_cols=len(header_list),
            display_values=self.expected.get(ws.title),
        )
        return ws

    # ============================================================ 任务2

    def _resolve_ratio_denominator(self, den_cfg):
        """
        解析任务2「占比」的分母：**「原始数据」两列合计之差**（合同金额合计 − 回款合计合计，
        即统计总览的「总合同金额」−「总回款合计」），并给出 pandas 期望值。

        配置来源：`task2_special_agg.ratio_denominator`
            {"minuend_field": "合同金额", "subtrahend_field": "回款合计"}

        两列合计经 `SUM('原始数据'!<列>3:<列>N)` 直接取自数据源，不引用统计总览。
        返回 `(minuend_col, subtrahend_col, expected_value)`。
        """
        cols, values = [], []
        for key in ("minuend_field", "subtrahend_field"):
            field = den_cfg[key]
            cols.append(self._col(field))
            values.append(float(AGG_SUM(self.df_raw[field])))
        return cols[0], cols[1], values[0] - values[1]

    def run_task2_formula_build(self, wb, field_to_col: dict):
        """任务2：未回款原因分类汇总。空的分类不视为一个分类；开启分类条件格式。"""
        ws = wb.create_sheet(title=self.t2_cfg["sheet_name"])
        group_field = self.t2_cfg["group_field"]
        metrics = self.t2_cfg["metrics"]
        header_list = (["序号", group_field]
                       + list(metrics.keys())
                       + [self.t2_cfg["ratio_field_name"]])
        ws.append([""])
        ws.append(header_list)

        group_col = self._col(group_field)
        pk_field = self.t2_cfg["contract_key"]
        pk_col = self._col(pk_field)
        amount_field = self.t2_cfg["amount_field"]
        # 「占比」分母 = 原始数据「合同金额」合计 −「回款合计」合计
        # （即统计总览的「总合同金额」−「总回款合计」；配置驱动 + pandas 期望值）
        den_min_col, den_sub_col, den_value = self._resolve_ratio_denominator(
            self.t2_cfg["ratio_denominator"])

        group_val_list = self.get_sorted_dim_list(group_field, self.t2_cfg["sort_by_field"])
        grouped = self.df_raw.groupby(group_field, sort=False)
        self.report["reason_category_count"] = len(group_val_list)

        # 循环外预计算每个指标的（源字段, 聚合类型, Excel 列字母），行循环内直接取用
        metric_plans = [(src_field, agg_type, self._col(src_field))
                        for src_field, agg_type in metrics.values()]

        # 「占比」列号（表尾第 1 列）与分子列号（第 3 列 = metrics 首项的金额求和列）
        # 均与行无关，循环外确定一次
        ratio_col = len(header_list)
        numerator_col_letter = get_column_letter(3)

        for i, g_val in enumerate(group_val_list):
            row = self.first_row + i
            g = grouped.get_group(g_val)
            ref = '"%s"' % g_val
            self._put(ws, row, 1, build_seq_formula(self.first_row - 1), i + 1)
            self._put(ws, row, 2, g_val, None)
            for j, (src_field, agg_type, src_col) in enumerate(metric_plans, start=3):
                if agg_type == "sum":
                    self._put(ws, row, j,
                              build_sumifs_formula(src_col, {group_col: ref},
                                                   first_row=self.first_row, last_row=self.last_row),
                              float(AGG_SUM(g[src_field])))
                elif agg_type == "nunique":
                    self._put(ws, row, j,
                              build_sumproud_distinct_contract_formula(
                                  group_col, ref, pk_col,
                                  first_row=self.first_row, last_row=self.last_row),
                              int(AGG_NUNIQUE(g[pk_field])))
                elif agg_type in ("count", "count_rows"):
                    # 涉及数据笔数 = 该分组的原始明细行数
                    self._put(ws, row, j,
                              build_countifs_formula({group_col: ref},
                                                     first_row=self.first_row,
                                                     last_row=self.last_row),
                              int(len(g)))
                else:
                    raise ValueError("【Task2】不支持的聚合策略：%s" % agg_type)
            # 占比 = 本类开票未回款 / (原始数据「合同金额」合计 −「回款合计」合计)
            # 分母口径由配置 task2_special_agg.ratio_denominator 驱动
            unpaid = float(AGG_SUM(g[amount_field]))
            self._put(ws, row, ratio_col,
                      build_divide_by_columns_diff_formula(
                          "%s%d" % (numerator_col_letter, row),
                          den_min_col, den_sub_col,
                          first_row=self.first_row, last_row=self.last_row),
                      (unpaid / den_value) if den_value else None)

        apply_sheet_format(
            ws, self.style_cfg,
            num_cols=self.t2_cfg["num_fields"],
            pct_cols=self.t2_cfg["pct_fields"],
            date_cols=[],
            cat_cols=self.t2_cfg["cat_fields"],
            int_cols=self.t2_cfg.get("int_fields", []),
            text_cols=self.t2_cfg.get("text_fields", []),
            header_row_idx=2,
            # 分类列是否叠加「差异化柔和背景条件格式」（题目正文要求），由配置开关控制：
            # false = 该列不设条件格式，与同行其他单元格一致、只走斑马纹（用户 2026-10-03 选定）。
            enable_cat_cond_format=self.style_cfg.get("reason_soft_cf", False),
            cat_field_name=group_field,
            sheet_title=self.sheet_title_cfg["task2"],
            total_title_cols=len(header_list),
            display_values=self.expected.get(ws.title),
        )
        return ws

    # ============================================================ 统计总览（最后一张）

    def run_overview_formula_build(self, wb, field_to_col: dict):
        """统计总览：8 项全局指标，全部动态公式；数字格式按操作类型决定。"""
        ws = wb.create_sheet(title=self.t3_cfg["overview_sheet_name"])
        ws.append([""])
        ws.append(["指标", "数值"])

        pk = self.t1_cfg["primary_key"]
        pk_col = self._col(pk)
        reason_field = self.t2_cfg["group_field"]
        reason_col = self._col(reason_field)
        serial_col = self._col(self.t1_cfg["new_serial_name"])

        # 「整体回款率」分子/分母的**源字段**由配置反查得到，零硬编码：
        # overview_metrics 中 agg_operator=="ratio" 的 numerator/denominator 是「输出字段名」，
        # 回到同一列表按 output_field 取其 source_field，即得「回款合计」「合同金额」。
        # 换一套业务字段（如订单金额 / 已结算金额）时，只改配置即可。
        _src_of = {it["output_field"]: it.get("source_field", "")
                   for it in self.t3_cfg["overview_metrics"]}
        _ratio_item = next((it for it in self.t3_cfg["overview_metrics"]
                            if it.get("agg_operator") == "ratio"), None)
        recv_field = _src_of.get(_ratio_item.get("numerator"), "") if _ratio_item else ""
        amt_field = _src_of.get(_ratio_item.get("denominator"), "") if _ratio_item else ""
        if not (recv_field and amt_field):
            # 兜底：任务2 声明的占比分母（minuend − subtrahend）
            _rd = self.t2_cfg.get("ratio_denominator") or {}
            amt_field = _rd.get("minuend_field", "")
            recv_field = _rd.get("subtrahend_field", "")

        total_amount = float(AGG_SUM(self.df_raw[amt_field]))
        total_received = float(AGG_SUM(self.df_raw[recv_field]))
        distinct_contracts = int(AGG_NUNIQUE(self.df_raw[pk]))
        non_blank_reasons = self.df_raw.loc[
            ~self.df_raw[reason_field].apply(is_blank), reason_field]
        reason_count = int(AGG_NUNIQUE(non_blank_reasons)) if len(non_blank_reasons) else 0

        cell_row_of = {}
        row_formats = []
        for i, item in enumerate(self.t3_cfg["overview_metrics"]):
            row = self.first_row + i
            out_name = item["output_field"]
            op = item["agg_operator"]
            src_field = item.get("source_field", "")
            cell_row_of[out_name] = row

            if op == "sum":
                formula = build_sum_all_formula(self._col(src_field),
                                                first_row=self.first_row, last_row=self.last_row)
                expected = float(AGG_SUM(self.df_raw[src_field]))
                fmt = self.style_cfg["amount_number_format"]
            elif op == "ratio":
                formula = build_divide_formula("B%d" % cell_row_of[item["numerator"]],
                                               "B%d" % cell_row_of[item["denominator"]])
                expected = (total_received / total_amount) if total_amount else None
                fmt = self.style_cfg["percent_number_format"]
            elif op == "nunique":
                if src_field == pk:
                    formula = build_distinct_all_formula(pk_col, first_row=self.first_row,
                                                         last_row=self.last_row)
                    expected = distinct_contracts
                else:
                    formula = build_distinct_nonblank_all_formula(
                        reason_col, first_row=self.first_row, last_row=self.last_row)
                    expected = reason_count
                fmt = self.style_cfg.get("int_number_format", "0")
            elif op == "count_rows":
                formula = build_counta_formula(serial_col, first_row=self.first_row,
                                               last_row=self.last_row)
                expected = int(len(self.df_raw))
                fmt = self.style_cfg.get("int_number_format", "0")
            else:
                raise ValueError("【总览】不支持的聚合操作：%s" % op)

            self._put(ws, row, 1, out_name, None)
            self._put(ws, row, 2, formula, expected)
            # 先落数字格式：列宽是按"显示内容"估的，格式晚设会导致估宽失真
            ws.cell(row, 2).number_format = fmt
            row_formats.append(fmt)

        apply_sheet_format(
            ws, self.style_cfg,
            num_cols=["数值"], pct_cols=[], int_cols=[], date_cols=[], cat_cols=[],
            text_cols=["指标"],
            header_row_idx=2,
            sheet_title=self.sheet_title_cfg["overview"],
            total_title_cols=2,
            set_number_format=False,   # 逐项金额/百分比格式已在上面设好，不能被覆盖
            display_values=self.expected.get(ws.title),
        )
        return ws

    # ============================================================ 任务3

    def run_task3_multi_dim_formula_build(self, wb, field_to_col: dict):
        """任务3：4 个维度统计表。每个维度 7 项指标（含合同数、回款率、未回款占比）。"""
        t3 = self.t3_cfg
        derived = t3["derived_ratio_list"]
        pk_field = self.t1_cfg["primary_key"]
        pk_col = self._col(pk_field)

        # ---- 循环外预计算 ----
        # 4 个维度表共用同一套 dim_metrics / derived_ratio_list，指标列与派生列的
        # 列位 / 列字母只取决于表头结构、与具体维度无关，全部算一次复用；
        # 「全局<指标>」分母的源字段与全表合计也不随维度/行变化，同样只算一次。
        metric_keys = list(t3["dim_metrics"].keys())
        metric_start_col = 3                       # 第 1 列序号、第 2 列维度值，指标从第 3 列起
        metric_plans = [
            (out_field, src_field, agg_type, self._col(src_field))
            for out_field, (src_field, agg_type) in t3["dim_metrics"].items()
        ]
        first_ratio_col = metric_start_col + len(metric_keys)   # 派生比率列起始列号（1 基）
        ratio_plans = []
        for k, ratio_cfg in enumerate(derived):
            num_field = ratio_cfg["numerator_field"]
            den_field = ratio_cfg["denominator_field"]
            plan = {
                "col": first_ratio_col + k,
                "num_field": num_field,
                "num_let": get_column_letter(metric_keys.index(num_field) + metric_start_col),
            }
            if str(den_field).startswith("全局"):
                # 「全局<指标>」= 全表该源字段合计；字段由配置反查，零硬编码
                base_field = self._resolve_global_base_field(den_field)
                plan["base_col"] = self._col(base_field)
                plan["base_total"] = float(AGG_SUM(self.df_raw[base_field]))
            else:
                plan["den_field"] = den_field
                plan["den_let"] = get_column_letter(
                    metric_keys.index(den_field) + metric_start_col)
            ratio_plans.append(plan)

        for dim_item in t3["dim_list"]:
            sheet_n = dim_item["sheet_name"]
            dim_field = dim_item["dimension_field"]
            dim_col = self._col(dim_field)
            ws = wb.create_sheet(title=sheet_n)

            header = (["序号", dim_field]
                      + metric_keys
                      + [x["output_field"] for x in derived])
            ws.append([""])
            ws.append(header)

            dim_val_list = self.get_sorted_dim_list(dim_field, t3.get("sort_by_field"))
            grouped = self.df_raw.groupby(dim_field, sort=False)
            self.report.setdefault("dim_category_counts", {})[dim_field] = len(dim_val_list)

            for i, d_val in enumerate(dim_val_list):
                row = self.first_row + i
                g = grouped.get_group(d_val)
                ref = '"%s"' % d_val
                self._put(ws, row, 1, build_seq_formula(self.first_row - 1), i + 1)
                self._put(ws, row, 2, d_val, None)

                # 各指标列：sum 记录组内合计（供派生比率作分子/分母），nunique 去重计数
                sums = {}
                for offset, (out_field, src_field, agg_type, src_col) in enumerate(metric_plans):
                    col = metric_start_col + offset
                    if agg_type == "sum":
                        val = float(AGG_SUM(g[src_field]))
                        sums[out_field] = val
                        self._put(ws, row, col,
                                  build_sumifs_formula(src_col, {dim_col: ref},
                                                       first_row=self.first_row,
                                                       last_row=self.last_row),
                                  val)
                    elif agg_type == "nunique":
                        self._put(ws, row, col,
                                  build_sumproud_distinct_contract_formula(
                                      dim_col, ref, pk_col,
                                      first_row=self.first_row, last_row=self.last_row),
                                  int(AGG_NUNIQUE(g[pk_field])))
                    else:
                        raise ValueError("【Task3】不支持的聚合策略：%s" % agg_type)

                # 派生比率：列位/列字母已在 ratio_plans 预计算，此处只做逐行求值
                for plan in ratio_plans:
                    num_val = sums.get(plan["num_field"])
                    if "base_col" in plan:            # 「全局<指标>」分母 = 全表合计
                        exp = (num_val / plan["base_total"]) if plan["base_total"] else None
                        formula = build_divide_by_global_formula(
                            "%s%d" % (plan["num_let"], row), plan["base_col"],
                            first_row=self.first_row, last_row=self.last_row)
                    else:                             # 同表内两列相除
                        den_val = sums.get(plan["den_field"])
                        exp = (num_val / den_val) if den_val else None
                        formula = build_divide_formula("%s%d" % (plan["num_let"], row),
                                                       "%s%d" % (plan["den_let"], row))
                    self._put(ws, row, plan["col"], formula, exp)

            apply_sheet_format(
                ws, self.style_cfg,
                num_cols=t3["num_fields"],
                pct_cols=t3["pct_fields"],
                date_cols=[],
                cat_cols=[dim_field],
                int_cols=t3.get("int_fields", []),
                text_cols=t3.get("text_fields", []),
                header_row_idx=2,
                sheet_title="%s%s" % (dim_field, self.sheet_title_cfg["dim_suffix"]),
                total_title_cols=len(header),
                display_values=self.expected.get(ws.title),
            )

    # ============================================================ 输出

    def _resolve_keep_sheets(self, wb, requested):
        """
        子集导出时的「公式依赖闭包」：被保留的工作表若在公式里引用了其它工作表，
        那些被引用的表也必须一起导出，否则公式会全部变成 `#REF!`。

        实现：先**单遍**扫描各表公式、建立 {表名: 被引用表集合} 的依赖图，
        再从请求保留的表出发做 BFS 扩散直至闭包——
        优于原先「每加一张表就重新全表扫描」的多轮写法。

        :param wb: 已构建完 8 张表的 workbook（尚未裁剪）
        :param requested: 调用方请求保留的表名列表
        :return: 按**原表序**排好的工作表名列表（保证「原始数据」仍在前）
        """
        all_names = [w.title for w in wb.worksheets]
        unknown = [s for s in requested if s not in all_names]
        if unknown:
            raise ValueError("【子集导出】找不到工作表 %s；可选：%s" % (unknown, all_names))

        # 单遍扫描建立依赖图：公式文本中出现「表名!」即视为引用了该表
        refs = {}
        for name in all_names:
            targets = set()
            for row in wb[name].iter_rows():
                for cell in row:
                    v = cell.value
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    targets.update(other for other in all_names
                                   if other != name and (other + "!") in v)
            refs[name] = targets

        # BFS 闭包：从请求表出发，把被引用表不断并入 keep，直到不再增长
        keep = list(dict.fromkeys(requested))
        queue = list(keep)
        while queue:
            for other in refs.get(queue.pop(), ()):
                if other not in keep:
                    keep.append(other)
                    queue.append(other)
        return [n for n in all_names if n in keep]

    def _apply_sheet_order(self, wb):
        """
        按配置 `profiles[<default_profile>].sheet_order` 重排工作表顺序。

        **构建顺序与交付顺序解耦**：改表序只改配置，不用动代码（单一事实源）。
        题目正文的编号顺序是「统计总览」排第 4 位（区域 / 部门 / 账龄 / 客户分类其后）。
        子集导出（--only）时只对保留下来的表排序；有未登记的表则直接报错，避免静默丢表。
        """
        want = list(self.config["profiles"][self.config["default_profile"]]["sheet_order"])
        actual = [w.title for w in wb.worksheets]
        missing = [n for n in actual if n not in want]
        if missing:
            raise ValueError("【输出】工作表 %s 未登记在配置 sheet_order=%s 中" % (missing, want))
        # openpyxl 无公开的重排 API，直接重排内部列表（社区通用做法）
        wb._sheets = [wb[name] for name in want if name in actual]

    def render_export_excel(self, output_file_path: str, only_sheets=None):
        """
        统一输出入口：
        1) 沿用源 workbook，仅把源工作表改名为「原始数据」——值 / 行列顺序 / 数字格式零改动；
        2) 依次生成 任务1 → 任务2 → 任务3 四维度 → 统计总览（构建顺序不影响交付顺序）；
        3) 按配置 `sheet_order` 重排为交付顺序（题目正文：统计总览第 4 位）；
        4) 保存后注入公式缓存值，并把运行事实写入 self.report（由 run_skill.py 落盘）。

        :param only_sheets: 只导出指定工作表（列表）——用于“单独出一个表”的场景。
            被保留的表若在公式里引用了别的表，那些表会**自动一并导出**（否则公式全成
            `#REF!`）；输出顺序仍按配置表序。`None` = 导出全部 8 张（默认交付物）。
        """
        t0 = time.time()
        wb = load_workbook(self.input_excel_path, data_only=False)
        raw_ws = wb[self.input_cfg["sheet_name"]]
        raw_ws.title = self.output_sheet_cfg["raw_copy"]

        header_row = self.input_cfg["header_row"]
        field_to_col = self.build_header_col_mapping(raw_ws, header_row)

        # 有效列数 = 从第 1 列起连续非空表头的个数（排除源文件自带的幽灵列）
        raw_cols = 0
        for c in range(1, raw_ws.max_column + 1):
            v = raw_ws.cell(header_row, c).value
            if v is not None and str(v).strip() != "":
                raw_cols += 1
            else:
                break

        # 「原始数据」零改动：默认**不叠加任何美化**（不改填充 / 字体 / 边框 / 对齐 /
        # 列宽 / 冻结），只做“改名为原始数据”这一步，确保与源表逐格一致
        # （题目注意事项(1)：不得修改原始数据的单元格格式）。
        # 若确需统一美化，把配置 style_setting.raw_beautify 设为 true。
        if self.style_cfg.get("raw_beautify", False):
            raw_fmt = self.config.get("raw_copy_format", {})
            apply_sheet_format(
                raw_ws, self.style_cfg,
                num_cols=raw_fmt.get("num_fields", self.t1_cfg["num_fields"]),
                pct_cols=[],
                date_cols=raw_fmt.get("date_fields", self.t1_cfg["date_fields"]),
                cat_cols=raw_fmt.get("cat_fields", self.t1_cfg["cat_fields"]),
                int_cols=raw_fmt.get("int_fields", []),
                text_cols=raw_fmt.get("text_fields", []),
                header_row_idx=header_row,
                sheet_title=None,            # 保留源第 1 行「单位：万元」，不覆盖为大标题
                total_title_cols=raw_cols,
                set_number_format=False,     # 数字格式逐格保持源文件原样
                display_values=self.expected.get(raw_ws.title),
            )

        self.run_task1_formula_build(wb, field_to_col)
        self.run_task2_formula_build(wb, field_to_col)
        self.run_task3_multi_dim_formula_build(wb, field_to_col)
        self.run_overview_formula_build(wb, field_to_col)

        # 子集导出（--only）：先做公式依赖闭包，再按配置表序裁剪掉其余工作表。
        # 不裁剪「原始数据」这类被引用表，避免公式变成 #REF!。
        if only_sheets:
            keep = self._resolve_keep_sheets(wb, list(only_sheets))
            for _name in [w.title for w in wb.worksheets]:
                if _name not in keep:
                    del wb[_name]
            expected = {k: v for k, v in self.expected.items() if k in keep}
        else:
            expected = self.expected

        # 交付顺序由配置决定（单一事实源）——构建顺序无关紧要
        self._apply_sheet_order(wb)

        try:
            wb.save(output_file_path)
        except PermissionError as exc:
            # 结果文件被 Excel / WPS / 网盘 / 杀毒软件占用时会走到这里。
            # 给出可执行的提示，而不是把 PermissionError 原样抛给调用方。
            raise RuntimeError(
                "输出文件被占用，无法写入：%s\n"
                "请先关闭 Excel / WPS 中打开的该文件（或等同步、杀毒扫描结束）后重试。"
                % output_file_path
            ) from exc
        finally:
            wb.close()

        # 注入公式缓存值，使程序化读取也能取到数值
        stats = inject_cached_values(output_file_path, expected)
        errors = scan_errors(output_file_path)

        # 运行回执：产物侧的客观事实。
        # 表序用只读模式重开产物读取，读完即关，避免文件句柄悬挂
        # （Windows 下句柄不释放会影响后续 verify.py / 网盘同步）。
        wb_check = load_workbook(output_file_path, read_only=True)
        sheet_order = [w.title for w in wb_check.worksheets]
        wb_check.close()
        self.report.update({
            "output": os.path.abspath(output_file_path),
            "sheet_order": sheet_order,
            # 回执必须**如实反映配置**（硬性规则 2 要求原始数据不得被美化）：
            # 此前这里是一句硬编码文案，raw_beautify=false 时仍写着"已叠加表头美化/斑马纹…"，属交付物事实错误。
            "raw_sheet": (
                "零美化：仅改表名；值/行列顺序/数字格式/填充·字体·边框·对齐/列宽/冻结全部与源表逐格一致"
                if not self.style_cfg.get("raw_beautify", False) else
                "值/行列顺序/数字格式原样保留，另叠加了表头美化、斑马纹、水平边框、列宽重算、冻结窗格"
            ),
            "formula_cells_written": stats["formula_cells"],
            "cached_values_injected": stats["injected"],
            "cached_by_sheet": stats["sheets"],
            "formula_errors": errors[:20],
            "formula_error_count": len(errors),
            "elapsed_sec": round(time.time() - t0, 2),
            "status": "success" if not errors else "success_with_formula_errors",
        })
        return output_file_path
