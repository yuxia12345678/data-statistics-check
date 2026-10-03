# -*- coding: utf-8 -*-
"""
Skill统一命令行启动入口 run_skill.py
支持命令行参数指定业务配置模板路径以及原始输入excel；
切换业务场景只需要更换--config指向不同的json模板；
调用示例：
python run_skill.py --config ../config/contract_repayment.json --input ./合同开票及回款核对表.xlsx
"""
import argparse
from stat_engine import GeneralStatSkillEngine


def main():
    """
    主入口函数：解析命令行参数；初始化引擎；加载数据源；执行全部统计并输出excel报表
    """
    parser = argparse.ArgumentParser(description="通用明细统计Skill运行入口，切换业务仅更换配置模板")
    parser.add_argument("--config", required=True, help="业务JSON配置模板文件路径")
    parser.add_argument("--input", required=True, help="原始输入excel路径")
    args = parser.parse_args()

    # 使用指定业务模板初始化统计引擎
    engine = GeneralStatSkillEngine(args.config)
    # 加载原始业务数据
    engine.load_source_data(args.input)
    # 读取配置中的输出文件名，执行渲染导出报表
    out_filename = engine.config["output_filename"]
    out_path = engine.render_export_excel(out_filename)
    print(f"【Skill执行结束】输出文件：{out_path}")


if __name__ == "__main__":
    main()

# 运行方式
#  python run_skill.py --config ../config/contract_repayment.json --input ./合同开票及回款核对表.xlsx
