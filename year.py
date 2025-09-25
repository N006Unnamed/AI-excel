import pandas as pd
import numpy as np
from openpyxl import load_workbook
import os

#     生成动态年份的excel表格


def generate_fund_table(input_file, start_year, end_year, output_file=None):
    """
    从Excel文件生成动态年份范围的流动资金估算表

    参数:
    input_file: 输入Excel文件路径
    start_year: 起始年份
    end_year: 结束年份
    output_file: 输出Excel文件路径(可选)

    返回:
    pandas DataFrame 结果表
    """
    # 验证年份范围
    if not (2025 <= start_year <= 2045) or not (2025 <= end_year <= 2045):
        raise ValueError("年份范围必须在2025-2045之间")
    if start_year > end_year:
        raise ValueError("起始年份不能大于结束年份")

    # 读取Excel文件
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"文件不存在: {input_file}")

    # 读取整个Excel文件
    df = pd.read_excel(input_file)

    # 获取所有年份列（假设年份列都是数字形式）
    year_cols = [col for col in df.columns if isinstance(col, (int, np.integer)) and 2025 <= col <= 2045]

    # 验证请求的年份范围
    if start_year not in year_cols or end_year not in year_cols:
        raise ValueError("请求的年份范围超出数据文件范围")

    # 生成年份序列
    years = [year for year in range(start_year, end_year + 1)]

    # 构建结果表
    # 固定列
    fixed_cols = ["序号", "项目", "最低周转天数(天)", "周转次数"]

    # 特殊处理建设期标识
    year_columns = []
    for year in years:
        if year == 2025:
            year_columns.append(f"{year} (建设期)")
        else:
            year_columns.append(str(year))

    # 选择需要的列
    result_df = df[fixed_cols + years].copy()

    # 重命名年份列
    result_df.columns = fixed_cols + year_columns

    # 如果需要输出到文件
    if output_file:
        result_df.to_excel(output_file, index=False)
        print(f"结果已保存到: {output_file}")

    return result_df


# 使用示例
if __name__ == "__main__":
    # 配置参数
    INPUT_FILE = "动态年份测试.xlsx"  # 原始数据文件
    OUTPUT_FILE = "流动资金估算表_动态生成.xlsx"  # 输出文件

    # 示例1: 生成2026-2028年的表格
    table_2026_2028 = generate_fund_table(
        input_file=INPUT_FILE,
        start_year=2026,
        end_year=2028,
        output_file=OUTPUT_FILE
    )

    # 打印结果
    print("2026-2028年流动资金估算表:")
    print(table_2026_2028.to_markdown(index=False))

    # 示例2: 生成2026-2030年的表格
    try:
        table_2026_2030 = generate_fund_table(
            input_file=INPUT_FILE,
            start_year=2026,
            end_year=2030,
            output_file="流动资金估算表_2026-2030.xlsx"
        )
    except ValueError as e:
        print(f"注意: {e} (需要确保原始数据包含2030年数据)")