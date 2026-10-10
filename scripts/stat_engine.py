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
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from agg_strategy import (
    AGG_STRATEGY_REGISTRY, distinct_join, first_non_null, is_blank,
    last_non_null, text_of,
)
from cached_values import inject_cached_values, scan_errors
from config_validator import kind_of_name, style_kind_of, validate_business_config
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
        self.first_row = self.input_cfg["header_row"] + 1     # 源表数据首行（公式区间用）
        # 产物工作表的数据首行**固定为 3**（第 1 行合并大标题 + 第 2 行表头），
        # 与源表表头行位置**解耦**：源表表头在第 3 行（其他行业报表常见）时，
        # 公式区间起点仍是源表数据首行，但产物数据行必须从第 3 行开始。
        self.out_first_row = 3
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

        # 动态探测数据末行。口径见 _resolve_last_row：默认仍是「最后一个非空行」，
        # 换「其他行业」源表（含「合计」行 / 页脚行）时可在配置 `input.data_end`
        # 里声明停止词或停止条件，**无需改代码**。
        self.last_row = self._resolve_last_row(ws, header_row)
        # 运行回执：源表侧的客观事实
        self.report.update({
            "input": os.path.abspath(input_excel_path),
            "config": os.path.abspath(self.config_path),
            "profile": self.config.get("default_profile", "spec"),
            "source_sheet": self.input_cfg["sheet_name"],
            "hidden_sheets": [w.title for w in wb.worksheets if w.sheet_state != "visible"],
            "header_row": header_row,
            "data_first_row": self.first_row,
            "data_last_row": self.last_row,
            "source_columns": len(self.field_to_col),
        })
        wb.close()

        # 只读取**数据区**（表头行 + 到 last_row 为止）。其他行业报表常在表尾带
        # 「合计」行与页脚行（如「内容由 AI 生成」），若不限定行数：
        #   1) 它们会被当成明细参与统计 —— 结果直接算错；
        #   2) 数值列会因混入 `=SUM(...)` 文本变成 object 类型，求和抛异常。
        df = pd.read_excel(input_excel_path,
                           sheet_name=self.input_cfg["sheet_name"],
                           skiprows=header_row - 1,
                           nrows=max(0, self.last_row - header_row))
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

    def _resolve_last_row(self, ws, header_row: int) -> int:
        """探测数据末行（口径由配置声明，换源表零改动）。

        优先级：
          1. `input.data_end.last_row`          —— 显式指定数据末行（最直接）；
          2. `input.data_end.stop_keywords`     —— 首列（去空格后）以这些词开头即停止
             （排除「合  计」「总计」「小计」这类汇总行）；
          3. `input.data_end.stop_on_blank_first_col = true`
             —— 首列为空即停止（排除表尾页脚行，如「内容由 AI 生成」）。
        以上都没配 → 保持原口径：最后一个含任意非空单元格的行。

        单遍 iter_rows 扫描（与旧实现同量级开销），不逐格 ws.cell 解析坐标。
        """
        cfg = self.input_cfg.get("data_end") or {}
        if cfg.get("last_row"):
            return int(cfg["last_row"])
        keywords = [str(k).replace(" ", "") for k in (cfg.get("stop_keywords") or [])]
        stop_blank = bool(cfg.get("stop_on_blank_first_col", False))
        last = header_row
        for row in ws.iter_rows(min_row=header_row + 1):
            if not keywords and not stop_blank:
                # 默认口径：最后一个含任意非空单元格的行
                if any(cell.value not in (None, "") for cell in row):
                    last = row[0].row
                continue
            raw = row[0].value
            text = "" if raw is None else str(raw).replace(" ", "").strip()
            if stop_blank and text == "":
                break
            if keywords and any(text.startswith(k) for k in keywords):
                break
            last = row[0].row
        return last

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
            row = self.out_first_row + i
            g = grouped.get_group(code)
            ref = '"%s"' % code
            self._put(ws, row, 1, build_seq_formula(self.out_first_row - 1), i + 1)
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

        # 「占比」列号（表尾第 1 列）与分子列号**均与行无关，循环外确定一次**。
        # 分子列 = metrics 中「来源字段 == amount_field」的那一列（配置反查，不写死列号）：
        # 增 / 删 / 重排 metrics、改业务字段名后，分子列自动跟随，代码零改动。
        ratio_col = len(header_list)
        numerator_col_letter = get_column_letter(3)
        for _j, (_src, _agg, _c) in enumerate(metric_plans, start=3):
            if _src == amount_field:
                numerator_col_letter = get_column_letter(_j)
                break

        for i, g_val in enumerate(group_val_list):
            row = self.out_first_row + i
            g = grouped.get_group(g_val)
            ref = '"%s"' % g_val
            self._put(ws, row, 1, build_seq_formula(self.out_first_row - 1), i + 1)
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

    def _overview_format(self, out_name: str, agg_operator: str) -> str:
        """统计总览某行的数字格式：先按 `field_kind` 反查该指标，再按聚合类型兜底。

        与 `verify.py` 的 R8 使用**同一个** `kind_of_name` 解析器（配置是唯一事实源），
        避免生成侧与自检侧对同一行给出不同期望。
        """
        kind = kind_of_name(self.config.get("field_kind") or {}, out_name)
        fmt_by_kind = {
            "money": self.style_cfg["amount_number_format"],
            "num": self.style_cfg["amount_number_format"],
            "int": self.style_cfg.get("int_number_format", "0"),
            "pct": self.style_cfg["percent_number_format"],
            "date": self.style_cfg.get("date_number_format", "@"),
        }
        if kind in fmt_by_kind:
            return fmt_by_kind[kind]
        return {
            "sum": self.style_cfg["amount_number_format"],
            "avg": self.style_cfg["amount_number_format"],
            "ratio": self.style_cfg["percent_number_format"],
            "nunique": self.style_cfg.get("int_number_format", "0"),
            "count_rows": self.style_cfg.get("int_number_format", "0"),
        }.get(agg_operator, self.style_cfg["amount_number_format"])

    def run_overview_formula_build(self, wb, field_to_col: dict):
        """统计总览：8 项全局指标，全部动态公式；数字格式按操作类型决定。"""
        ws = wb.create_sheet(title=self.t3_cfg["overview_sheet_name"])
        ws.append([""])
        ws.append(["指标", "数值"])

        pk = self.t1_cfg["primary_key"]
        pk_col = self._col(pk)
        reason_field = self.t2_cfg["group_field"]
        reason_col = self._col(reason_field)
        # 「总数据笔数」的计数列：`task3_multi_dim.count_rows_field` 优先 → 业务主键列
        # → 源表第 1 个业务列。**不能写死取「序号」列**：其他行业源表常常没有序号列
        # （序号是合并表的输出列，不是源字段），否则会直接报「找不到字段」。
        count_field = (self.t3_cfg.get("count_rows_field")
                       or (pk if pk in self.field_to_col else "")
                       or next(iter(self.field_to_col), ""))
        serial_col = self._col(count_field)

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
            row = self.out_first_row + i
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
                # 去重计数的**列按配置的 source_field 反查**（原先非主键一律落到
                # 任务2 分类列，等于写死了一种业务字段，增删字段时会算错）。
                if src_field == pk:
                    formula = build_distinct_all_formula(pk_col, first_row=self.first_row,
                                                         last_row=self.last_row)
                    expected = distinct_contracts
                elif src_field == reason_field:
                    formula = build_distinct_nonblank_all_formula(
                        reason_col, first_row=self.first_row, last_row=self.last_row)
                    expected = reason_count
                else:
                    _col_letter = self._col(src_field)
                    _nb = self.df_raw.loc[
                        ~self.df_raw[src_field].apply(is_blank), src_field]
                    formula = build_distinct_nonblank_all_formula(
                        _col_letter, first_row=self.first_row, last_row=self.last_row)
                    expected = int(AGG_NUNIQUE(_nb)) if len(_nb) else 0
                fmt = self.style_cfg.get("int_number_format", "0")
            elif op == "count_rows":
                formula = build_counta_formula(serial_col, first_row=self.first_row,
                                               last_row=self.last_row)
                expected = int(len(self.df_raw))
                fmt = self.style_cfg.get("int_number_format", "0")
            else:
                raise ValueError("【总览】不支持的聚合操作：%s" % op)

            # 数字格式：**优先按 field_kind 反查该指标**（与其它表同一套规则），
            # 查不到再按聚合类型兜底。这样给总览新增非金额指标（如「总销售数量（只）」
            # 登记为 int）只需改配置，不必改代码，且与自检 R8 的判据完全一致。
            fmt = self._overview_format(out_name, op)

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
                row = self.out_first_row + i
                g = grouped.get_group(d_val)
                ref = '"%s"' % d_val
                self._put(ws, row, 1, build_seq_formula(self.out_first_row - 1), i + 1)
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

    # ============================================================ 通用声明式工作表
    #
    # 目的（对应《竞赛须知》「规则可配置」）：**换行业数据 / 新增工作表 /
    # 表内加·删·改字段，只改配置 JSON，本文件零改动。**
    # 配置形态与完整语义见 references/配置驱动规范.md（§2 声明式 columns、§3 extra_sheets）。
    #
    # 支持的列聚合 `agg`：
    #   group          分组值本身（分类/文本列，写字面量）
    #   seq            序号 =ROW()-n
    #   const          固定文本（literal）
    #   sum            SUMIFS(原始数据!<field>, 原始数据!<分组列>, <本行分组值单元格>)
    #   count          COUNTIFS(原始数据!<分组列>, <本行分组值单元格>) = 组内明细行数
    #   nunique        SUMPRODUCT((分组列=值)/COUNTIFS(域,域&"",分组列,分组列&""))
    #   first_non_null INDEX+MATCH 取组内第一条非空
    #   last_non_null  同 first_non_null（取末条；公式同 first 构造器的区间语义）
    #   distinct_join  组内非空去重拼接（Excel 无非数组稳定写法 → 公式化例外，写文本）
    #   ratio          分子/分母 = 本表输出列名 | "global:<源字段>" | "literal:<数>"；
    #                  分母亦可为 {"minuend_field":..,"subtrahend_field":..}（两列合计之差）
    # 合计行（total_row=true）：group 列写 total_label；sum/count/nunique 写全表聚合公式；
    #   ratio 引用合计行自身的单元格；seq/const/文本类留空。

    def _col_letter(self, column_index: int) -> str:
        return get_column_letter(column_index)

    def _group_keys(self, group_field: str) -> list:
        """分组取值清单：[(分组键文本, 原始值)]，按首次出现顺序，空白不计入。"""
        seen, out = set(), []
        for v in self.df_raw[group_field].tolist():
            k = text_of(v)
            if not k or k in seen:
                continue
            seen.add(k)
            out.append((k, v))
        return out

    def _subset(self, group_field, group_key):
        """该分组的明细行；group_key 为 None → 全表（合计行 / 全表汇总行用）。"""
        if group_field is None or group_key is None:
            return self.df_raw
        return self.df_raw[self.df_raw[group_field].map(
            lambda v: text_of(v) == group_key)]

    def _eval_column(self, col: dict, g):
        """非 ratio 列的期望值（pandas 独立算，仅供缓存值注入 / 自检复核）。"""
        agg = col["agg"]
        if agg == "sum":
            return float(AGG_SUM(g[col["field"]]))
        if agg == "nunique":
            nb = g.loc[~g[col["field"]].apply(is_blank), col["field"]]
            return int(AGG_NUNIQUE(nb)) if len(nb) else 0
        if agg == "count":
            return int(len(g))
        if agg == "first_non_null":
            v = first_non_null(g[col["field"]])
            return v if not is_blank(v) else ""
        if agg == "last_non_null":
            v = last_non_null(g[col["field"]])
            return v if not is_blank(v) else ""
        if agg == "distinct_join":
            return distinct_join(g[col["field"]], col.get("join_separator", "；"))
        if agg == "const":
            return col.get("literal", "")
        return None                                  # group / seq / ratio 由调用方处理

    def _formula_column(self, col: dict, row: int, ctx: dict):
        """某列在第 row 行的 Excel 公式；返回 None = 该格写字面量（分组值 / 文本拼接）。"""
        agg = col["agg"]
        if agg in ("group", "const", "distinct_join"):
            return None
        if agg == "seq":
            return build_seq_formula(self.out_first_row - 1)
        raw, r0, r1 = ctx["raw_sheet"], self.first_row, self.last_row
        cond = {}
        if ctx["group_col"] and ctx["cond_cell"]:
            cond = {ctx["group_col"]: ctx["cond_cell"] % row}
        if agg in ("sum", "count", "nunique", "first_non_null", "last_non_null"):
            fcol = self._col(col["field"])
            if not cond:                             # 合计行 / 全表汇总行 → 全表聚合公式
                if agg == "sum":
                    return build_sum_all_formula(fcol, source_sheet=raw,
                                                 first_row=r0, last_row=r1)
                if agg == "count":
                    return build_counta_formula(ctx["group_col"], source_sheet=raw,
                                                first_row=r0, last_row=r1)
                if agg == "nunique":
                    return build_distinct_nonblank_all_formula(fcol, source_sheet=raw,
                                                               first_row=r0, last_row=r1)
                return None                          # 文本类没有「全表」公式，留空
            if agg == "sum":
                return build_sumifs_formula(fcol, cond, source_sheet=raw,
                                            first_row=r0, last_row=r1)
            if agg == "count":
                return build_countifs_formula(cond, source_sheet=raw,
                                              first_row=r0, last_row=r1)
            if agg == "nunique":
                return build_sumproud_distinct_contract_formula(
                    ctx["group_col"], cond[ctx["group_col"]], fcol,
                    source_sheet=raw, first_row=r0, last_row=r1)
            return build_index_first_non_null(ctx["group_col"], cond[ctx["group_col"]],
                                              fcol, source_sheet=raw,
                                              first_row=r0, last_row=r1)
        if agg == "ratio":
            num = self._operand(col["numerator"], row, ctx)
            den_cfg = col["denominator"]
            if isinstance(den_cfg, dict):
                return build_divide_by_columns_diff_formula(
                    num, self._col(den_cfg["minuend_field"]),
                    self._col(den_cfg["subtrahend_field"]),
                    source_sheet=raw, first_row=r0, last_row=r1)
            return build_divide_formula(num, self._operand(den_cfg, row, ctx))
        raise ValueError("【extra_sheets】不支持的聚合：%s" % agg)

    def _operand(self, ref, row: int, ctx: dict) -> str:
        """ratio 的分子 / 分母：本表输出列名 → 单元格；global:<源字段> → 全表合计。"""
        if isinstance(ref, (int, float)) and not isinstance(ref, bool):
            return repr(float(ref))
        s = str(ref)
        if s.startswith("global:"):
            return build_sum_all_formula(self._col(s.split(":", 1)[1]),
                                         source_sheet=ctx["raw_sheet"],
                                         first_row=self.first_row,
                                         last_row=self.last_row)[1:]
        if s.startswith("literal:"):
            return s.split(":", 1)[1]
        cell = ctx["cells"].get(s)
        if not cell:
            raise ValueError("ratio 引用的列「%s」未在本表 columns 中声明" % s)
        return cell % row

    def _ratio_value(self, col: dict, get_value):
        """ratio 列的期望值（pandas 独立算）。"""
        def operand(ref):
            if isinstance(ref, (int, float)) and not isinstance(ref, bool):
                return float(ref)
            s = str(ref)
            if s.startswith("global:"):
                return float(AGG_SUM(self.df_raw[s.split(":", 1)[1]]))
            if s.startswith("literal:"):
                return float(s.split(":", 1)[1])
            return get_value(s)

        den_cfg = col["denominator"]
        if isinstance(den_cfg, dict):
            den = (float(AGG_SUM(self.df_raw[den_cfg["minuend_field"]]))
                   - float(AGG_SUM(self.df_raw[den_cfg["subtrahend_field"]])))
        else:
            den = operand(den_cfg)
        num = operand(col["numerator"])
        try:
            num, den = float(num), float(den)
        except (TypeError, ValueError):
            return None
        return (num / den) if den else None

    @staticmethod
    def _kinds_of(cols: list) -> dict:
        """按列声明的 kind 汇总成 apply_sheet_format 需要的字段类型清单
        （money → num：金额在样式清单里叫 num，在 field_kind 里叫 money）。"""
        out = {"num": [], "int": [], "pct": [], "date": [], "cat": [], "text": []}
        for c in cols:
            out.setdefault(style_kind_of(c["kind"]), []).append(c["name"])
        return out

    def run_extra_sheets_build(self, wb, field_to_col):
        """构建配置 `extra_sheets` 声明的全部新增工作表（type=group / pivot）。"""
        for spec in (self.config.get("extra_sheets") or []):
            if spec.get("type", "group") == "pivot":
                self._build_pivot_sheet(wb, spec)
            else:
                self._build_group_sheet(wb, spec)
            self.report.setdefault("extra_sheets", []).append(spec["sheet_name"])

    # ---------------------------------------------------------- group 型（逐组一行）
    def _build_group_sheet(self, wb, spec):
        ws = wb.create_sheet(title=spec["sheet_name"])
        cols = list(spec["columns"])
        serial = spec.get("serial")
        header = ([serial] if serial else []) + [c["name"] for c in cols]
        ws.append([""])
        ws.append(header)

        ctx = {"raw_sheet": self.output_sheet_cfg["raw_copy"],
               "group_field": spec.get("group_field"),
               "group_col": (self._col(spec["group_field"])
                             if spec.get("group_field") else None),
               "cond_cell": None,
               "cells": {}}
        group_out_idx = None
        for i, c in enumerate(cols, start=(2 if serial else 1)):
            ctx["cells"][c["name"]] = self._col_letter(i) + "%d"
            if group_out_idx is None and c["agg"] == "group":
                group_out_idx = i
        if group_out_idx:
            ctx["cond_cell"] = "$%s%%d" % self._col_letter(group_out_idx)

        keys = (self._group_keys(spec["group_field"]) if spec.get("group_field")
                else [(None, "")])
        rows_data = []
        for key, raw_val in keys:
            g = self._subset(spec.get("group_field"), key)
            vals = {}
            for c in cols:
                if c["agg"] == "group":
                    vals[c["name"]] = raw_val
                elif c["agg"] not in ("ratio", "seq"):
                    vals[c["name"]] = self._eval_column(c, g)
            rows_data.append({"key": key, "raw_val": raw_val, "g": g, "vals": vals})
        # ratio 第二遍：可引用同表任意非 ratio 列（列序不受限）
        for rd in rows_data:
            for c in cols:
                if c["agg"] == "ratio":
                    rd["vals"][c["name"]] = self._ratio_value(
                        c, lambda n, rd=rd: rd["vals"].get(n))

        sort_by = spec.get("sort_by")
        if sort_by:
            def _sort_key(rd):
                try:
                    return float(rd["vals"].get(sort_by))
                except (TypeError, ValueError):
                    return float("-inf")
            rows_data.sort(key=_sort_key, reverse=bool(spec.get("sort_desc", True)))

        for i, rd in enumerate(rows_data):
            row = self.out_first_row + i
            c_idx = 1
            if serial:
                self._put(ws, row, c_idx, build_seq_formula(self.out_first_row - 1), i + 1)
                c_idx += 1
            for c in cols:
                exp = (i + 1) if c["agg"] == "seq" else rd["vals"].get(c["name"])
                formula = self._formula_column(c, row, ctx)
                if formula is None:
                    val = rd["raw_val"] if c["agg"] == "group" else exp
                    self._put(ws, row, c_idx, val, val)
                else:
                    self._put(ws, row, c_idx, formula,
                              exp if exp is not None else "")
                c_idx += 1

        if spec.get("total_row"):
            self._append_total_row(ws, spec, cols, serial, ctx, header)

        self._format_declared_sheet(ws, spec, cols, serial, self._kinds_of(cols))
        return ws

    def _append_total_row(self, ws, spec, cols, serial, ctx, header):
        """合计行：数值列写全表聚合公式，文本列留空，分组列写 total_label。"""
        n_rows = ws.max_row - 2                       # 已有数据行数（不含标题/表头）
        trow = self.out_first_row + n_rows
        label = spec.get("total_label", "合计")
        total_vals = {}
        for c in cols:
            agg = c["agg"]
            if agg == "sum":
                total_vals[c["name"]] = float(AGG_SUM(self.df_raw[c["field"]]))
            elif agg == "nunique":
                total_vals[c["name"]] = int(AGG_NUNIQUE(self.df_raw[c["field"]]))
            elif agg == "count":
                total_vals[c["name"]] = (
                    int(AGG_SUM(self.df_raw[ctx["group_field"]].notna()))
                    if ctx["group_field"] else int(len(self.df_raw)))
        for c in cols:
            if c["agg"] == "ratio":
                total_vals[c["name"]] = self._ratio_value(c, lambda n: total_vals.get(n))

        ctx_total = dict(ctx)
        ctx_total["cond_cell"] = None                 # 无分组条件 → 全表聚合公式
        c_idx = 2 if serial else 1
        for c in cols:
            agg = c["agg"]
            if agg == "group":
                val, formula = label, label
            elif agg in ("sum", "count", "nunique", "ratio"):
                val = total_vals.get(c["name"])
                formula = self._formula_column(c, trow, ctx_total)
            else:
                val, formula = "", None
            if formula:
                self._put(ws, trow, c_idx, formula, val if val is not None else "")
            elif val != "":
                self._put(ws, trow, c_idx, val, val)
            c_idx += 1
        for cc in range(1, len(header) + 1):
            ws.cell(trow, cc).font = Font(bold=True)

    def _format_declared_sheet(self, ws, spec, cols, serial, kinds):
        """统一美化声明式工作表（样式参数仍全部来自 style_setting）。"""
        int_cols = list(kinds.get("int", []))
        if serial:
            int_cols.append(serial)
        apply_sheet_format(
            ws, self.style_cfg,
            num_cols=list(kinds.get("num", [])),
            pct_cols=list(kinds.get("pct", [])),
            date_cols=list(kinds.get("date", [])),
            cat_cols=list(kinds.get("cat", [])),
            int_cols=int_cols,
            text_cols=list(kinds.get("text", [])),
            header_row_idx=2,
            sheet_title=spec.get("sheet_title") or spec["sheet_name"],
            total_title_cols=ws.max_column,
            display_values=self.expected.get(ws.title),
        )

    # ---------------------------------------------------------- pivot 型（行列交叉）
    def _pivot_value(self, row_field, col_field, val_field, agg, row_key, col_key):
        """交叉表单元格期望值（pandas 独立算）。"""
        df = self.df_raw
        if row_key is not None:
            df = df[df[row_field].map(lambda v: text_of(v) == row_key)]
        if col_key is not None:
            df = df[df[col_field].map(lambda v: text_of(v) == col_key)]
        if agg == "count":
            return int(len(df))
        return float(AGG_SUM(df[val_field]))

    @staticmethod
    def _pivot_formula(agg, vcol, rcol, ccol, row_ref, col_ref, raw, r0, r1):
        cond = {}
        if row_ref:
            cond[rcol] = row_ref
        if col_ref:
            cond[ccol] = col_ref
        if agg == "count":
            return build_countifs_formula(cond, source_sheet=raw, first_row=r0, last_row=r1)
        return build_sumifs_formula(vcol, cond, source_sheet=raw, first_row=r0, last_row=r1)

    def _build_pivot_sheet(self, wb, spec):
        ws = wb.create_sheet(title=spec["sheet_name"])
        row_field, col_field = spec["row_field"], spec["col_field"]
        val_field = spec["value_field"]
        agg = spec.get("agg", "sum")
        label = spec.get("total_label", "合计")
        total_col = bool(spec.get("total_col", True))
        total_row = bool(spec.get("total_row", True))

        raw, r0, r1 = self.output_sheet_cfg["raw_copy"], self.first_row, self.last_row
        vcol, rcol, ccol = self._col(val_field), self._col(row_field), self._col(col_field)

        row_keys = self._group_keys(row_field)
        col_keys = self._group_keys(col_field)
        if spec.get("col_sort") == "desc":            # 按列合计降序（默认保持首次出现顺序）
            col_keys.sort(key=lambda kv: -self._pivot_value(
                row_field, col_field, val_field, agg, None, kv[0]))

        header = ([row_field] + [c for c, _v in col_keys]
                  + ([label] if total_col else []))
        ws.append([""])
        ws.append(header)

        for i, (rk, rv) in enumerate(row_keys):
            row = self.out_first_row + i
            ws.cell(row, 1).value = rv
            for j, (ck, _cv) in enumerate(col_keys):
                col_ref = "%s$2" % self._col_letter(2 + j)
                self._put(ws, row, 2 + j,
                          self._pivot_formula(agg, vcol, rcol, ccol, "$A%d" % row,
                                              col_ref, raw, r0, r1),
                          self._pivot_value(row_field, col_field, val_field, agg, rk, ck))
            if total_col:                             # 行合计：只按行条件求和（仍引用数据源）
                self._put(ws, row, 2 + len(col_keys),
                          self._pivot_formula(agg, vcol, rcol, ccol, "$A%d" % row,
                                              None, raw, r0, r1),
                          self._pivot_value(row_field, col_field, val_field, agg, rk, None))

        if total_row:
            trow = self.out_first_row + len(row_keys)
            ws.cell(trow, 1).value = label
            for j, (ck, _cv) in enumerate(col_keys):
                self._put(ws, trow, 2 + j,
                          self._pivot_formula(agg, vcol, rcol, ccol, None,
                                              "%s$2" % self._col_letter(2 + j),
                                              raw, r0, r1),
                          self._pivot_value(row_field, col_field, val_field, agg, None, ck))
            if total_col:
                if agg == "count":
                    f = build_counta_formula(rcol, source_sheet=raw,
                                             first_row=r0, last_row=r1)
                else:
                    f = build_sum_all_formula(vcol, source_sheet=raw,
                                              first_row=r0, last_row=r1)
                self._put(ws, trow, 2 + len(col_keys), f,
                          self._pivot_value(row_field, col_field, val_field, agg, None, None))
            for cc in range(1, len(header) + 1):
                ws.cell(trow, cc).font = Font(bold=True)

        apply_sheet_format(
            ws, self.style_cfg,
            num_cols=[c for c, _v in col_keys] + ([label] if total_col else []),
            pct_cols=[], int_cols=[], date_cols=[],
            cat_cols=[row_field], text_cols=[],
            header_row_idx=2,
            sheet_title=spec.get("sheet_title") or spec["sheet_name"],
            total_title_cols=len(header),
            display_values=self.expected.get(ws.title),
        )
        return ws

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

        # 「原始数据」可能自带公式（报表尾部的「合计」行等）。openpyxl 保存时会丢掉
        # 公式的**缓存值**，程序化读取（data_only=True）遂读到空，自检 R2a 会判为
        # 与源表不一致。这里把源文件里已有的缓存值原样登记进 expected，
        # 保存后由 cached_values 写回 <v>，保证「原始数据」逐格与源表一致。
        _cache_wb = load_workbook(self.input_excel_path, data_only=True)
        _cache_ws = _cache_wb[self.input_cfg["sheet_name"]]
        for _row in raw_ws.iter_rows():
            for _cell in _row:
                if isinstance(_cell.value, str) and _cell.value.startswith("="):
                    _cached = _cache_ws.cell(_cell.row, _cell.column).value
                    if _cached is not None:
                        self.expected.setdefault(raw_ws.title, {})[
                            _cell.coordinate] = _cached
        _cache_wb.close()

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
        # 配置 `extra_sheets` 声明的**新增工作表**（group / pivot 两种类型），
        # 表名与字段全部来自 JSON —— 加表 / 表内加·删·改字段都不需要改本文件。
        self.run_extra_sheets_build(wb, field_to_col)

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
