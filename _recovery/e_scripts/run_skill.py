# -*- coding: utf-8 -*-
"""
Skill命令行统一入口 run_skill.py
调用示例：
python run_skill.py --config ../config/contract_repayment.json --input ./合同开票及回款核对表.xlsx
输出文件命名：合同开票及回款核对表分析结果_姓名.xlsx
特性：配置驱动，更换config即可适配其他业务规则，代码无需改动。
"""
import argparse
import traceback
from stat_engine import GeneralStatSkillEngine


def main():
    parser = argparse.ArgumentParser(description="合同数据统计核对Skill（动态Excel公式输出）")
    parser.add_argument("--config", required=True, help="业务JSON配置文件路径")
    parser.add_argument("--input", required=True, help="输入原始Excel路径：合同开票及回款核对表.xlsx")
    args = parser.parse_args()
    try:
        engine = GeneralStatSkillEngine(args.config)
        engine.load_source_data(args.input)
        output_filename = engine.config["output_filename"]
        out_path = engine.render_export_excel(output_filename)
        print(f"✅ Skill执行完成，输出文件：{out_path}")
    except Exception as e:
        print(f"❌ Skill执行异常：{str(e)}")
        traceback.print_exc()
        exit(1)


if __name__ == "__main__":
    main()
