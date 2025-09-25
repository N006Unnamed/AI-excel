import pandas as pd
import numpy as np
import os
import sys


def calculate_fund_table(input_file, start_year, end_year, output_file=None):
    """
    从Excel文件生成动态年份范围的流动资金估算表，所有值基于公式计算

    参数:
    input_file: 输入Excel文件路径
    start_year: 起始年份
    end_year: 结束年份
    output_file: 输出Excel文件路径(可选)

    返回:
    pandas DataFrame 结果表
    """
    # 验证年份范围
    if start_year > end_year:
        raise ValueError("起始年份不能大于结束年份")
    if start_year < 2025 or end_year > 2050:
        raise ValueError("年份范围必须在2025-2050之间")

    # 读取Excel文件
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"文件不存在: {input_file}")

    # 读取整个Excel文件
    df = pd.read_excel(input_file)

    # 获取所有年份列
    year_cols = [col for col in df.columns if isinstance(col, (int, np.integer)) and 2025 <= col <= 2045]

    # 添加新列用于扩展年份
    for year in range(start_year, end_year + 1):
        if year not in df.columns and year > 2045:
            df[year] = np.nan

    # 创建序号到行的映射
    index_map = {}
    for idx, row in df.iterrows():
        index_map[row['序号']] = idx

    # 计算所有年份的值（基于固定公式）
    for year in range(start_year, end_year + 1):
        # 跳过已有数据的年份
        if year in year_cols:
            continue

        # 计算每个项目的值（基于固定公式）
        for idx, row in df.iterrows():
            item = row['项目']

            # 跳过特殊行（在产品）
            if item == "在产品":
                df.at[idx, year] = "根据实际判断是否有"
                continue

            # 根据项目类型应用不同公式
            if item == "原材料":
                # 原材料 = 上年值 * (1 + 年增长率)
                base_value = df.at[idx, 2045]
                df.at[idx, year] = base_value * (1 + 0.01) ** (year - 2045)

            elif item == "燃料动力":
                # 燃料动力 = 上年值 * (1 + 年增长率)
                base_value = df.at[idx, 2045]
                df.at[idx, year] = base_value * (1 + 0.01) ** (year - 2045)

            elif item == "产成品":
                # 产成品 = (应收账款 + 原材料 + 燃料动力) * 系数
                ar_idx = index_map['1.1']
                material_idx = index_map['1.2.1']
                fuel_idx = index_map['1.2.2']

                ar_value = df.at[ar_idx, year] if not pd.isna(df.at[ar_idx, year]) else df.at[ar_idx, 2045]
                material_value = df.at[material_idx, year] if not pd.isna(df.at[material_idx, year]) else df.at[
                    material_idx, 2045]
                fuel_value = df.at[fuel_idx, year] if not pd.isna(df.at[fuel_idx, year]) else df.at[fuel_idx, 2045]

                df.at[idx, year] = (ar_value + material_value + fuel_value) * 0.25

            elif item == "现金":
                # 现金 = (应收账款 + 产成品) * 系数
                ar_idx = index_map['1.1']
                product_idx = index_map['1.2.4']

                ar_value = df.at[ar_idx, year] if not pd.isna(df.at[ar_idx, year]) else df.at[ar_idx, 2045]
                product_value = df.at[product_idx, year] if not pd.isna(df.at[product_idx, year]) else df.at[
                    product_idx, 2045]

                df.at[idx, year] = (ar_value + product_value) * 0.15

            elif item == "应收账款":
                # 应收账款 = 上年值 * (1 + 年增长率)
                base_value = df.at[idx, 2045]
                df.at[idx, year] = base_value * (1 + 0.01) ** (year - 2045)

            elif item == "应付账款":
                # 应付账款 = (原材料 + 燃料动力) * 系数
                material_idx = index_map['1.2.1']
                fuel_idx = index_map['1.2.2']

                material_value = df.at[material_idx, year] if not pd.isna(df.at[material_idx, year]) else df.at[
                    material_idx, 2045]
                fuel_value = df.at[fuel_idx, year] if not pd.isna(df.at[fuel_idx, year]) else df.at[fuel_idx, 2045]

                df.at[idx, year] = (material_value + fuel_value) * 0.05

            elif item == "存货":
                # 存货 = 原材料 + 燃料动力 + 产成品
                material_idx = index_map['1.2.1']
                fuel_idx = index_map['1.2.2']
                product_idx = index_map['1.2.4']

                material_value = df.at[material_idx, year] if not pd.isna(df.at[material_idx, year]) else df.at[
                    material_idx, 2045]
                fuel_value = df.at[fuel_idx, year] if not pd.isna(df.at[fuel_idx, year]) else df.at[fuel_idx, 2045]
                product_value = df.at[product_idx, year] if not pd.isna(df.at[product_idx, year]) else df.at[
                    product_idx, 2045]

                df.at[idx, year] = material_value + fuel_value + product_value

            elif item == "流动资产":
                # 流动资产 = 应收账款 + 存货 + 现金
                ar_idx = index_map['1.1']
                inventory_idx = index_map['1.2']
                cash_idx = index_map['1.3']

                ar_value = df.at[ar_idx, year] if not pd.isna(df.at[ar_idx, year]) else df.at[ar_idx, 2045]
                inventory_value = df.at[inventory_idx, year] if not pd.isna(df.at[inventory_idx, year]) else df.at[
                    inventory_idx, 2045]
                cash_value = df.at[cash_idx, year] if not pd.isna(df.at[cash_idx, year]) else df.at[cash_idx, 2045]

                df.at[idx, year] = ar_value + inventory_value + cash_value

            elif item == "流动负债":
                # 流动负债 = 应付账款
                ap_idx = index_map['2.1']
                ap_value = df.at[ap_idx, year] if not pd.isna(df.at[ap_idx, year]) else df.at[ap_idx, 2045]
                df.at[idx, year] = ap_value

            elif item == "流动资金（1-2）":
                # 流动资金 = 流动资产 - 流动负债
                assets_idx = index_map['1']
                liabilities_idx = index_map['2']

                assets_value = df.at[assets_idx, year] if not pd.isna(df.at[assets_idx, year]) else df.at[
                    assets_idx, 2045]
                liabilities_value = df.at[liabilities_idx, year] if not pd.isna(df.at[liabilities_idx, year]) else \
                df.at[liabilities_idx, 2045]

                df.at[idx, year] = assets_value - liabilities_value

            elif item == "流动资金当年增加额":
                # 增加额 = 当年流动资金 - 上年流动资金
                wc_idx = index_map['3']
                current_wc = df.at[wc_idx, year] if not pd.isna(df.at[wc_idx, year]) else df.at[wc_idx, 2045]

                # 获取上年流动资金
                prev_year = year - 1
                if prev_year in df.columns:
                    prev_wc = df.at[wc_idx, prev_year] if not pd.isna(df.at[wc_idx, prev_year]) else 0
                else:
                    prev_wc = df.at[wc_idx, 2045]  # 如果没有上年数据，使用2045年数据

                df.at[idx, year] = current_wc - prev_wc

            elif item == "流动资金贷款":
                # 2028年后流动资金贷款为0
                if year > 2028:
                    df.at[idx, year] = 0
                else:
                    # 2028年前使用2045年数据
                    df.at[idx, year] = df.at[idx, 2045]

    # 生成年份序列
    years = [year for year in range(start_year, end_year + 1)]

    # 构建结果表
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

    # 处理特殊值行（在产品）
    special_rows = result_df["项目"] == "在产品"
    result_df.loc[special_rows, year_columns] = "根据实际判断是否有"

    # 处理特殊值行（流动资金贷款）
    loan_rows = result_df["项目"] == "流动资金贷款"
    for year in years:
        if year > 2028:
            col_name = str(year) if year != 2025 else f"{year} (建设期)"
            result_df.loc[loan_rows, col_name] = 0

    # 如果需要输出到文件
    if output_file:
        result_df.to_excel(output_file, index=False)
        print(f"结果已保存到: {output_file}")

    return result_df


def get_valid_year(prompt, min_year=2025, max_year=2050):
    """
    获取有效的年份输入
    """
    while True:
        try:
            year = int(input(prompt))
            if min_year <= year <= max_year:
                return year
            else:
                print(f"年份必须在{min_year}-{max_year}范围内，请重新输入。")
        except ValueError:
            print("请输入有效的年份数字。")


def main():
    print("流动资金估算表生成工具 (支持2025-2050年)")
    print("=" * 50)

    # 配置输入文件
    INPUT_FILE = "动态年份测试.xlsx"

    # 检查文件是否存在
    if not os.path.exists(INPUT_FILE):
        print(f"错误: 数据文件 '{INPUT_FILE}' 不存在")
        print("请确保数据文件在当前目录下")
        sys.exit(1)

    # 用户输入年份范围
    print("\n请指定要生成的年份范围 (2025-2050)")
    start_year = get_valid_year("请输入起始年份: ", 2025, 2050)
    end_year = get_valid_year("请输入结束年份: ", 2025, 2050)

    if start_year > end_year:
        print("起始年份不能大于结束年份，已自动交换年份")
        start_year, end_year = end_year, start_year

    # 生成输出文件名
    output_file = f"流动资金估算表_{start_year}-{end_year}.xlsx"

    # 生成表格
    try:
        print(f"\n正在生成 {start_year}-{end_year} 年的流动资金估算表...")
        table = calculate_fund_table(
            input_file=INPUT_FILE,
            start_year=start_year,
            end_year=end_year,
            output_file=output_file
        )

        print("\n生成成功!")
        print(f"表格已保存为: {output_file}")

        # 显示预览
        print("\n表格预览 (前5行):")
        print(table.head().to_markdown(index=False))

        # 打开文件选项
        open_file = input("\n是否要打开生成的Excel文件? (y/n): ").lower()
        if open_file == 'y':
            try:
                os.startfile(output_file)  # Windows
            except:
                try:
                    os.system(f'open "{output_file}"')  # macOS
                except:
                    os.system(f'xdg-open "{output_file}"')  # Linux

    except Exception as e:
        print(f"\n生成表格时出错: {str(e)}")


if __name__ == "__main__":
    main()