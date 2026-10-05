# -*- coding: utf-8 -*-
"""
Skill 命令行统一入口 run_skill.py

调用示例：
    python run_skill.py --config config/contract_repayment.json \
                        --input "输入/附件1-合同开票及回款核对表.xlsx" \
                        --output-dir 运行成果

输出文件命名：`合同开票及回款核对表分析结果_<姓名>.xlsx`（文件名取自配置 output_filename）。

产物策略（本次新增）
--------------------
1. **只对「姓名」发问**：交付物文件名里的姓名若不是真实姓名（占位符「姓名」
   或为空），本脚本会**交互式询问一次**；用户回车即采用配置默认值。
   其余任何信息（口径、样式、路径……）都不再追问，一律按配置执行。
2. **回执/自检报告覆盖式更新**：`运行回执.json`、`自检报告.json` 每次运行
   覆盖上一次（`"w"` 模式），目录里始终只留最新一份。
3. **输入 / 输出只保留最新一次结果**：详见引擎 `output.keep_latest_only` 配置项
   —— 每次运行会清理「输入」目录里的旧源表、以及**实际输出目录**里的旧结果文件，
   只保留本次使用/产出的那一份。

特性：配置驱动，更换 config 即可适配其他业务规则，代码无需改动。
"""
import argparse
import json
import os
import sys
import traceback

RECEIPT_NAME = "运行回执.json"

# 需要向用户确认的唯一一件事：交付物文件名里的「姓名」。
# 配置里若仍是这个占位符（或为空），说明还没填真实姓名，才需要发问。
PLACEHOLDER_NAMES = {"姓名", ""}

# Windows 控制台/管道可能是 GBK 编码，直接 print 非 GBK 字符（如 ✅/❌）会抛
# UnicodeEncodeError，把一次成功的运行误报成失败。把 stdout 的错误策略改为 replace，
# 保证任何输出环境（控制台、管道、重定向、被工具捕获）都不会因编码中断。
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

# ---------------------------------------------------------------- 运行前依赖自检
# 必须放在 `import stat_engine`（进而 import pandas）之前：
#   依赖齐全 → 跳过安装，直接继续执行；
#   依赖缺失 → 优先用工程内 vendor/ 离线安装，成功后再继续。
# 实现见 scripts/deps_check.py。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deps_check import DependencyError, ensure_dependencies

try:
    ensure_dependencies()
except DependencyError as exc:
    print("❌ 运行前依赖自检未通过，已终止执行：\n      %s" % exc)
    sys.exit(1)

from stat_engine import GeneralStatSkillEngine


def _split_filename(filename):
    """把 output_filename 拆成 (前缀, 姓名, 扩展名)。

    只认最后一个下划线后的、扩展名之前的部分为「姓名」段：
    `合同开票及回款核对表分析结果_张三.xlsx` → ("合同开票及回款核对表分析结果", "张三", ".xlsx")
    无下划线时把整个主名当作姓名段。
    """
    stem, ext = os.path.splitext(os.path.basename(filename))
    if "_" in stem:
        prefix, name = stem.rsplit("_", 1)
        return prefix, name, ext
    return "", stem, ext


def _resolve_name(output_filename, no_prompt=False):
    """解析交付物文件名里的「姓名」，必要时交互式询问一次（唯一的用户交互点）。

    - 姓名段已有真实值（不是占位符）→ 直接返回，不打扰用户；
    - 姓名段是占位符 / 空 → 询问一次；用户输入有效值则采用，否则用配置默认值；
    - `--no-name-prompt` 或非交互式环境（无 stdin）→ 不问，直接用配置值。
    """
    prefix, name, ext = _split_filename(output_filename)
    if name and name not in PLACEHOLDER_NAMES:
        return name, False          # 已是真实姓名，无需询问

    if no_prompt or not sys.stdin or not sys.stdin.isatty():
        return name, False          # 非交互环境不阻塞，保持配置值

    try:
        ans = input("请输入交付物文件名中的【姓名】（直接回车则沿用配置值「%s」）：" % name)
    except (EOFError, KeyboardInterrupt):
        print()                     # 用户中断输入：回落配置值，不视为错误
        return name, False
    ans = (ans or "").strip()
    if ans and ans not in PLACEHOLDER_NAMES:
        print("   已将交付物姓名改为「%s」" % ans)
        return ans, True
    return name, False


def main():
    parser = argparse.ArgumentParser(description="合同数据统计核对Skill（动态Excel公式输出）")
    parser.add_argument("--config", required=True, help="业务JSON配置文件路径")
    parser.add_argument("--input", required=True, help="输入原始Excel路径：合同开票及回款核对表.xlsx")
    parser.add_argument("--output-dir", default=None,
                        help="输出目录；不传则写到当前工作目录（输出文件名取自配置 output_filename）")
    parser.add_argument("--only", default=None,
                        help="只导出指定工作表（逗号分隔，如 --only 同合同号合并汇总）；"
                             "公式引用到的表（原始数据）会自动一并导出，避免 #REF!；"
                             "文件名会追加 _<表名>，且不写运行回执（避免覆盖完整交付物那份）")
    parser.add_argument("--report", default=None,
                        help="运行回执 JSON 路径；不传则写到结果文件同目录的 运行回执.json")
    parser.add_argument("--name", default=None,
                        help="直接指定交付物文件名里的「姓名」，指定后不再交互询问")
    parser.add_argument("--no-name-prompt", action="store_true",
                        help="禁止交互询问姓名，一律沿用配置 output_filename 里的值")
    args = parser.parse_args()

    try:
        only = [s.strip() for s in args.only.split(",") if s.strip()] if args.only else None
        engine = GeneralStatSkillEngine(args.config)
        engine.load_source_data(args.input)

        # ---- 唯一的用户交互点：确定交付物姓名 ----
        output_filename = engine.config["output_filename"]
        prefix, name, ext = _split_filename(output_filename)
        if args.name:                       # 命令行显式指定优先，且不再询问
            name = args.name.strip()
        else:
            name, _changed = _resolve_name(output_filename,
                                           no_prompt=args.no_name_prompt)
        base_name = "%s_%s%s" % (prefix, name, ext) if prefix else "%s%s" % (name, ext)

        if only:
            # 子集导出：文件名追加 _<表名>，不动完整交付物
            stem, ext2 = os.path.splitext(base_name)
            base_name = "%s_%s%s" % (stem, "_".join(only), ext2)

        if args.output_dir:
            os.makedirs(args.output_dir, exist_ok=True)
            out_path = os.path.join(args.output_dir, base_name)
        else:
            out_path = base_name
        out_path = engine.render_export_excel(out_path, only_sheets=only)

        # ---- 清理：输入 / 输出目录只保留本次最新结果 ----
        cleaned = _cleanup_latest(engine, args.input, out_path)
        for line in cleaned:
            print("   %s" % line)

        if only:
            # 子集导出的回执只描述这份子集文件；若与完整交付物同目录会把它那份覆掉，故不写。
            print(f"✅ Skill执行完成（子集导出），输出文件：{out_path}")
            print("   已跳过运行回执（避免覆盖完整交付物的 运行回执.json）")
            return

        # 运行回执：与结果文件同目录落盘（"w" 覆盖式），保证"结果 + 回执 + 自检报告"三份产物齐全
        report_path = args.report or os.path.join(os.path.dirname(out_path) or ".", RECEIPT_NAME)
        report_dir = os.path.dirname(report_path)
        if report_dir:
            os.makedirs(report_dir, exist_ok=True)
        with open(report_path, "w", encoding="utf-8") as fh:
            json.dump(engine.report, fh, ensure_ascii=False, indent=2, default=str)

        print(f"✅ Skill执行完成，输出文件：{out_path}")
        print(f"   运行回执：{report_path}")
    except Exception as e:
        print(f"❌ Skill执行异常：{str(e)}")
        traceback.print_exc()
        exit(1)


# ---------------------------------------------------------------- 产物清理（只留最新）
def _cleanup_latest(engine, input_path, out_path):
    """按配置 `output.keep_latest_only` 清理输入 / 输出目录，只保留本次产物。

    安全性（务必保持）：
      • 只删除**明确识别为上一次产物**的文件——输入目录里同扩展名的旧源表、
        输出目录里同扩展名的旧结果 / 回执；
      • 绝不删除系统/个人目录，绝不递归，绝不使用通配符批量删；
      • 本次使用的源文件与本次产出的结果文件永不删除（显式保护）；
      • 每一项删除都会打印出来，失败只告警不中断。
    """
    output_cfg = engine.config.get("output") or {}
    if not output_cfg.get("keep_latest_only", False):
        return []

    notes = []
    keep = {os.path.abspath(p) for p in (input_path, out_path)}

    def _sweep(directory, keep_exts, keep_names, label):
        if not directory or not os.path.isdir(directory):
            return
        for fn in sorted(os.listdir(directory)):
            fp = os.path.join(directory, fn)
            if not os.path.isfile(fp):
                continue
            ap = os.path.abspath(fp)
            if ap in keep or fn in keep_names:
                continue
            ext = os.path.splitext(fn)[1].lower()
            if ext not in keep_exts:
                continue
            try:
                os.remove(fp)
                notes.append("已清理旧%s：%s" % (label, fn))
            except OSError as exc:
                notes.append("⚠️ 旧%s未能删除（%s）：%s" % (label, fn, exc))

    # 输入目录：源表一般为 .xlsx / .xls（保留本次使用的那一份）
    _sweep(output_cfg.get("input_dir"), {".xlsx", ".xls"},
           {os.path.basename(input_path)}, "输入文件")
    # 输出目录：清「实际产出目录」（out_path 所在目录），而非配置 output_dir——
    # --output-dir 指到别处时，绝不能误清默认目录（如 运行成果/）里刚生成的回执/自检报告
    _sweep(os.path.dirname(os.path.abspath(out_path)), {".xlsx", ".xls", ".json"},
           {os.path.basename(out_path)}, "输出文件")
    return notes


if __name__ == "__main__":
    main()
