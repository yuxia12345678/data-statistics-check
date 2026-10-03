# -*- coding: utf-8 -*-
"""
Skill命令行统一入口 run_skill.py
调用示例：
python run_skill.py --config ../config/contract_repayment.json --input ./合同开票及回款核对表.xlsx
输出文件命名：合同开票及回款核对表分析结果_姓名.xlsx
特性：配置驱动，更换config即可适配其他业务规则，代码无需改动。
"""
import argparse
import json
import os
import sys
import traceback
from stat_engine import GeneralStatSkillEngine

RECEIPT_NAME = "运行回执.json"

# Windows 控制台/管道可能是 GBK 编码，直接 print 非 GBK 字符（如 ✅/❌）会抛
# UnicodeEncodeError，把一次成功的运行误报成失败。把 stdout 的错误策略改为 replace，
# 保证任何输出环境（控制台、管道、重定向、被工具捕获）都不会因编码中断。
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass


def main():
    parser = argparse.ArgumentParser(description="合同数据统计核对Skill（动态Excel公式输出）")
    parser.add_argument("--config", required=True, help="业务JSON配置文件路径")
    parser.add_argument("--input", required=True, help="输入原始Excel路径：合同开票及回款核对表.xlsx")
    parser.add_argument("--output-dir", default=None,
                        help="输出目录；不传则写到当前工作目录（输出文件名取自配置 output_filename）")
    parser.add_argument("--report", default=None,
                        help="运行回执 JSON 路径；不传则写到结果文件同目录的 运行回执.json")
    args = parser.parse_args()
    try:
        engine = GeneralStatSkillEngine(args.config)
        engine.load_source_data(args.input)
        output_filename = engine.config["output_filename"]
        if args.output_dir:
            os.makedirs(args.output_dir, exist_ok=True)
            out_path = os.path.join(args.output_dir, os.path.basename(output_filename))
        else:
            out_path = output_filename
        out_path = engine.render_export_excel(out_path)

        # 运行回执：与结果文件同目录落盘，保证"结果 + 回执 + 自检报告"三份产物齐全
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


if __name__ == "__main__":
    main()
