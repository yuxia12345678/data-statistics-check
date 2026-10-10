# -*- coding: utf-8 -*-
"""
配置校验模块 config_validator
职责：
1、加载JSON业务配置后，运行业务逻辑前做schema校验；
2、检查必填节点、字段，提前抛出异常，避免中途运行报错；
3、严格校验工作表名称与顺序（**表数量亦取自配置声明**），与配置期望值不匹配直接抛出ValueError；
4、只校验配置结构，不校验业务数据。

可配置性说明（对应《竞赛须知》"Skill规则可配置要求"）
----------------------------------------------------
本模块**不写死任何业务表名**：期望的表名与顺序取自
`profiles.<default_profile>.sheet_order`，与实际输出表名（`output_sheet_names` +
`task2_special_agg.sheet_name` + `task3_multi_dim.overview_sheet_name` + `dim_list`）
双向核对。因此换一套业务配置 = 换一套表名与规则，代码零改动。
**引擎内不再保留任何业务表名兜底值**：配置未声明 `sheet_order` 时直接报错，
    而不是回落到内置常量——兜底常量会把本题表名带进通用引擎，破坏「代码零业务字面量」。
"""
from typing import Dict, Any

import re

# 期望的工作表数量**不写死**：唯一事实源 = 配置 `profiles.<default_profile>.sheet_order`
# 的项数。竞赛题目要求 8 张，即在 `config/contract_repayment.json` 里声明为 8 项；
# 换成别的表数（如增删维度表）只需改配置，本模块与引擎均零改动。


def expected_sheet_order(cfg: Dict[str, Any]):
    """返回期望的工作表**名称与顺序**（数量 = 该清单项数，不写死）。

    唯一事实源：`profiles.<default_profile>.sheet_order`。
    配置未声明时**直接抛错**：内置兜底常量会把本题业务表名注入通用引擎，
    换业务时会被误当作期望值，因此一律要求配置显式声明。
    """
    profiles = cfg.get("profiles") or {}
    prof = profiles.get(cfg.get("default_profile")) if cfg.get("default_profile") else None
    if not isinstance(prof, dict):
        prof = next((p for p in profiles.values() if isinstance(p, dict)), None) or {}
    order = prof.get("sheet_order")
    if not order:
        raise ValueError(
            "业务配置缺少 `profiles.<default_profile>.sheet_order`（工作表名称与顺序清单）。"
            "该清单必须由业务配置显式声明——引擎不内置任何业务表名兜底值，"
            "否则换业务后本题表名会被误当作期望值。"
        )
    return list(order)


def validate_business_config(cfg: Dict[str, Any]) -> None:
    """
    业务配置整体校验入口
    :param cfg: json.load读取后的配置字典
    :return: 无返回，校验失败抛ValueError
    """
    # 顶层必填节点
    must_top_keys = [
        "input", "style_setting", "output_filename", "output_sheet_names",
        "task1_group_merge", "task2_special_agg", "task3_multi_dim"
    ]
    for k in must_top_keys:
        if k not in cfg:
            raise ValueError(f"配置缺失顶层节点：{k}")

    input_cfg = cfg["input"]
    for k in ["sheet_name", "header_row"]:
        if k not in input_cfg:
            raise ValueError(f"input节点缺失字段：{k}")

    # 任务1校验
    t1 = cfg["task1_group_merge"]
    if "primary_key" not in t1 or "agg_strategy_list" not in t1:
        raise ValueError("task1_group_merge 需要配置 primary_key、agg_strategy_list")

    # 任务2校验
    t2 = cfg["task2_special_agg"]
    for k in ["group_field", "amount_field", "contract_key", "sheet_name",
              "metrics", "target_sum_field", "ratio_field_name", "sort_by_field"]:
        if k not in t2:
            raise ValueError(f"task2_special_agg缺失配置项：{k}")

    # 任务3校验
    t3 = cfg["task3_multi_dim"]
    for k in ["dim_list", "overview_sheet_name", "dim_metrics",
              "derived_ratio_list", "sort_by_field", "overview_metrics"]:
        if k not in t3:
            raise ValueError(f"task3_multi_dim缺失配置项：{k}")

    # 工作表名校验：期望值来自配置（profiles.<default_profile>.sheet_order），
    # 而非写死在代码里 —— 换案例/换规则时随配置一起替换，代码零改动。
    out_sheet = cfg["output_sheet_names"]
    for k in ["raw_copy", "task1_result"]:
        if k not in out_sheet:
            raise ValueError(f"output_sheet_names缺失配置项：{k}")

    # 声明式扩展：extra_sheets 可追加任意多张工作表（表名与字段全部来自配置）
    extra_names = extra_sheet_names(cfg)
    validate_extra_sheets(cfg)
    validate_declarative_columns(cfg)

    canonical = [
        out_sheet["raw_copy"],
        out_sheet["task1_result"],
        t2["sheet_name"],
        t3["overview_sheet_name"]
    ]
    for dim_item in t3["dim_list"]:
        canonical.append(dim_item["sheet_name"])
    sheet_collect = canonical + extra_names

    expect_sheet_order = expected_sheet_order(cfg)
    expect_count = len(expect_sheet_order)   # 表数量取自配置声明，代码不写死
    if len(sheet_collect) != expect_count or len(set(sheet_collect)) != expect_count:
        raise ValueError(
            f"工作表数量或重名错误！期望 {expect_count} 张互不重名的表"
            f"（期望数量取自配置 sheet_order），实际:{sheet_collect}"
        )
    if set(sheet_collect) != set(expect_sheet_order):
        missing = [s for s in sheet_collect if s not in expect_sheet_order]
        extra = [s for s in expect_sheet_order if s not in sheet_collect]
        raise ValueError(
            "工作表名称错误！名字、文字、大小写必须与构建结果完全一致\n"
            f"  声明:{sheet_collect}\n"
            f"  sheet_order:{expect_sheet_order}\n"
            f"  缺失:{missing} 多余:{extra}"
        )
    # 顺序：剔除 extra_sheets 后，其余表在 sheet_order 里的**相对顺序**必须与构建顺序
    # 一致（extra_sheets 允许插在任意位置）；无 extra_sheets 时则要求逐项完全一致。
    if extra_names:
        ordered_core = [s for s in expect_sheet_order if s not in extra_names]
        if ordered_core != canonical:
            raise ValueError(
                "工作表顺序错误！剔除 extra_sheets 后，其余表的先后顺序必须与构建顺序一致\n"
                f"  期望:{canonical}\n"
                f"  实际:{ordered_core}"
            )
    elif expect_sheet_order != canonical:
        missing = [s for s in expect_sheet_order if s not in sheet_collect]
        extra = [s for s in sheet_collect if s not in expect_sheet_order]
        raise ValueError(
            "工作表名称/顺序错误！名字、文字、大小写、先后顺序必须完全一致\n"
            f"  期望:{expect_sheet_order}\n"
            f"  实际:{sheet_collect}\n"
            f"  缺失:{missing} 多余:{extra}"
        )


# ================================================================ 声明式 columns / extra_sheets

# 声明式列支持的聚合（与 stat_engine 通用列引擎一一对应）
COLUMN_AGG_SUPPORTED = (
    "group", "seq", "const", "sum", "count", "nunique",
    "first_non_null", "last_non_null", "distinct_join", "ratio",
)
# 声明式新增工作表支持的类型
EXTRA_SHEET_TYPES = ("group", "pivot")
# 需要 field（源字段名）的聚合
COLUMN_AGG_NEEDS_FIELD = ("sum", "nunique", "first_non_null", "last_non_null",
                          "distinct_join")
# `field_kind` 用 money 表示金额（供自检的 R8/R9 判判格式与对齐）；
# `apply_sheet_format` 的样式清单则用 num（右对齐 + 金额格式）。两者等价，
# 由本别名统一，避免声明式列的 kind 在两侧“各说各话”。
KIND_ALIASES = {"money": "num"}


def style_kind_of(kind: str) -> str:
    """把声明式列的 kind 归一到 `apply_sheet_format` 的样式类型。"""
    return KIND_ALIASES.get(kind, kind)


# 表头名末尾的说明性括号，如「总合同金额（万元）」→「总合同金额」
_PAREN_SUFFIX_RE = re.compile(r"[（(][^）)]*[）)]$")


def kind_of_name(field_kind: Dict[str, Any], name):
    """按**表头名**反查 `field_kind` 的类型（生成侧与自检侧共用，确保两侧一致）。

    依次尝试：原名 → 去掉末尾括号说明（「（万元）」「（元）」等）→ 再去掉「总」前缀。
    这样「总合同金额（万元）」能反查到 `field_kind["总合同金额"] = money`，
    新增指标只需登记一个表头名，不必在每个表头里重复登记全称。
    查不到返回 None（调用方自行兜底）。
    """
    if name is None:
        return None
    head = str(name)
    base = _PAREN_SUFFIX_RE.sub("", head)
    for cand in (head, base):
        if cand in field_kind:
            return field_kind[cand]
        if cand.startswith("总") and cand[1:] in field_kind:
            return field_kind[cand[1:]]
    return None


def extra_sheet_names(cfg: Dict[str, Any]) -> list:
    """`extra_sheets` 声明的全部工作表名（顺序即声明顺序）。"""
    return [s.get("sheet_name") for s in (cfg.get("extra_sheets") or [])]


def declared_columns(cfg: Dict[str, Any]) -> list:
    """产出声明式 columns 清单：[(出处标签, [列对象, ...]), ...]。

    仅 **extra_sheets（新增工作表）** 使用声明式 `columns`。
    **既有工作表**的字段增/删/改同样只改配置，用的是它们各自的既有清单键：
      • 合并表      `task1_group_merge.agg_strategy_list`（可增删改聚合列）
      • 分类汇总表  `task2_special_agg.metrics` + `ratio_field_name`
      • 维度表      `task3_multi_dim.dim_metrics` + `derived_ratio_list`
      • 统计总览    `task3_multi_dim.overview_metrics`
    再叠加驱动格式与对齐的 `*_fields` 与 `field_kind`。这些全是 JSON 里的明文清单，
    因此「表内加·删·改字段」无需改代码。
    """
    out = []
    for i, spec in enumerate(cfg.get("extra_sheets") or []):
        if spec.get("columns"):
            out.append(("extra_sheets[%d].columns" % i, spec["columns"]))
    return out


def validate_declarative_columns(cfg: Dict[str, Any]) -> None:
    """校验声明式列清单的结构，并强制与 `field_kind` 登记一致。

    为什么要与 `field_kind` 双向校验：数字格式与对齐由 `field_kind` 驱动，
    `verify.py` 的 R8/R9 据此断言。若列声明里的 kind 与 `field_kind` 不一致，
    生成侧与自检侧就会各说各话。因此**要求两边显式登记且一致**，改字段时
    两处一起改（都在同一份 JSON 里），仍然零代码改动。
    """
    field_kind = cfg.get("field_kind") or {}
    for label, cols in declared_columns(cfg):
        names = []
        for j, col in enumerate(cols):
            if not isinstance(col, dict):
                raise ValueError("%s[%d] 必须是对象（含 name / kind / agg）" % (label, j))
            name = col.get("name")
            if not name:
                raise ValueError("%s[%d] 缺少 name（输出列名）" % (label, j))
            agg = col.get("agg")
            if agg not in COLUMN_AGG_SUPPORTED:
                raise ValueError("%s「%s」的 agg=%r 不支持；可选：%s"
                                 % (label, name, agg, "、".join(COLUMN_AGG_SUPPORTED)))
            if agg in COLUMN_AGG_NEEDS_FIELD and not col.get("field"):
                raise ValueError("%s「%s」agg=%s 需要 field（源表字段名）"
                                 % (label, name, agg))
            if agg == "ratio" and (col.get("numerator") is None
                                   or col.get("denominator") is None):
                raise ValueError("%s「%s」agg=ratio 需要 numerator 与 denominator"
                                 % (label, name))
            kind = col.get("kind") or field_kind.get(name)
            if not kind:
                raise ValueError(
                    "%s「%s」未声明 kind，且 field_kind 中也没登记该列名。\n"
                    "数字格式与对齐由 field_kind 驱动（自检 R8/R9 据此断言），"
                    "请在 field_kind 中登记该列。" % (label, name))
            if name in field_kind and style_kind_of(field_kind[name]) != style_kind_of(kind):
                raise ValueError(
                    "%s「%s」的 kind=%s 与 field_kind 登记的 %s 不一致；"
                    "两处必须一致（自检以 field_kind 为准）。"
                    % (label, name, kind, field_kind[name]))
            names.append(name)
        if len(set(names)) != len(names):
            dup = [n for n in names if names.count(n) > 1]
            raise ValueError("%s 列名重复：%s" % (label, sorted(set(dup))))


def validate_extra_sheets(cfg: Dict[str, Any]) -> None:
    """校验 `extra_sheets`（新增工作表）声明。"""
    names = []
    for i, spec in enumerate(cfg.get("extra_sheets") or []):
        where = "extra_sheets[%d]" % i
        if not isinstance(spec, dict):
            raise ValueError("%s 必须是对象" % where)
        name = spec.get("sheet_name")
        if not name:
            raise ValueError("%s 缺少 sheet_name" % where)
        names.append(name)
        stype = spec.get("type", "group")
        if stype not in EXTRA_SHEET_TYPES:
            raise ValueError("%s.type=%r 不支持；可选：%s"
                             % (where, stype, "、".join(EXTRA_SHEET_TYPES)))
        if stype == "group":
            if not spec.get("columns"):
                raise ValueError("%s（type=group）需要 columns" % where)
            if "group_field" not in spec:
                raise ValueError("%s（type=group）需要显式声明 group_field"
                                 "（写 null = 全表汇总一行）" % where)
            if spec.get("sort_by"):
                declared = [c.get("name") for c in spec["columns"]]
                if spec["sort_by"] not in declared:
                    raise ValueError("%s.sort_by=「%s」不在 columns 的输出列中（可选：%s）"
                                     % (where, spec["sort_by"], declared))
        else:
            for k in ("row_field", "col_field", "value_field"):
                if not spec.get(k):
                    raise ValueError("%s（type=pivot）缺少 %s" % (where, k))
            if spec.get("agg", "sum") not in ("sum", "count"):
                raise ValueError("%s（type=pivot）的 agg 仅支持 sum / count" % where)
    if len(set(names)) != len(names):
        raise ValueError("extra_sheets 的 sheet_name 重复：%s" % names)


# ================================================================ 字段类型清单校验

# 驱动对齐与数字格式的字段类型清单键（与 excel_styler.apply_sheet_format 一一对应）
FIELD_KIND_KEYS = [
    "num_fields", "int_fields", "pct_fields", "date_fields", "cat_fields", "text_fields"
]


def validate_sheet_field_kinds(sheet_name: str,
                               headers,
                               kind_lists: Dict[str, Any],
                               kind_keys=None) -> None:
    """
    校验一张工作表的**表头是否全部登记了字段类型**。

    对齐规则（文本左 / 数值右 / 日期与分类居中）与数字格式都由这些清单驱动：
    num_fields / int_fields / pct_fields / date_fields / cat_fields / text_fields。

    若有表头既不在上述任何清单中，说明该列类型未登记——它会被静默地兜底成文本
    （左对齐、不设数字格式），很可能不是预期结果。因此这里**直接抛 ValueError**，
    强制配置方显式登记每一列。

    :param sheet_name: 工作表名（仅用于错误信息）
    :param headers:    表头名称序列（None / 空字符串会被忽略）
    :param kind_lists: 形如 {"num_fields": [...], "text_fields": [...], ...}
    :param kind_keys:  要检查的清单键；默认 FIELD_KIND_KEYS
    :return: 无返回；存在未登记表头时抛 ValueError
    """
    keys = list(kind_keys) if kind_keys else list(FIELD_KIND_KEYS)
    classified = set()
    for k in keys:
        classified.update(kind_lists.get(k) or [])
    unknown = [h for h in headers if h not in (None, "") and h not in classified]
    if unknown:
        raise ValueError(
            "工作表「%s」有表头未登记字段类型：%s\n"
            "请在配置的 %s 中登记这些列（纯文本列请放入 text_fields）。"
            % (sheet_name, unknown, " / ".join(keys))
        )
