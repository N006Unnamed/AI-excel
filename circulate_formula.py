import os
import re
from openpyxl_vba import load_workbook

#  根据输入的excel表，选定数据来源表A、B、C等，通过公式计算例如：AxB+C，将公式写入目标sheet选定的位置

class ExcelFormulaGenerator:
    def __init__(self, data_file, output_file):
        """
        :param data_file: 包含所有数据的工作簿路径
        :param output_file: 输出文件路径（与输入文件相同）
        """
        self.data_file = data_file
        self.output_file = output_file

        # 加载工作簿（保留公式）
        self.wb = load_workbook(data_file)
        print(f"工作簿加载成功! 工作表: {', '.join(self.wb.sheetnames)}")

    def get_formula_reference(self, sheet_name, cell_ref):
        """生成工作表单元格的引用公式"""
        # 如果 sheet_name 为空字符串，说明是本表引用
        if not sheet_name or sheet_name == "":
            return cell_ref

        # 检查工作表名称是否包含需要引号包裹的特殊字符
        # 包括: 空格、中文、括号、连字符等
        needs_quotes = any(char in sheet_name for char in ' ()-（）【】[]') or not sheet_name.isascii()

        if needs_quotes:
            # 如果工作表名称中包含单引号，需要转义
            escaped_sheet_name = sheet_name.replace("'", "''")
            return f"'{escaped_sheet_name}'!{cell_ref}"
        return f"{sheet_name}!{cell_ref}"

    def shift_cell(self, cell_ref, col_shift=0, row_shift=0):
        """将单元格引用向右/向下移动指定位置，支持区域引用"""

        # 处理区域引用（如"F4:F24"）
        if ':' in cell_ref:
            start_ref, end_ref = cell_ref.split(':')
            shifted_start = self._shift_single_cell(start_ref, col_shift, row_shift)
            shifted_end = self._shift_single_cell(end_ref, col_shift, row_shift)
            return f"{shifted_start}:{shifted_end}"
        else:
            return self._shift_single_cell(cell_ref, col_shift, row_shift)

    def _shift_single_cell(self, cell_ref, col_shift, row_shift):
        """处理单个单元格的偏移，支持绝对地址（$）"""

        # 判断是否包含绝对引用符号 $
        col_abs = '$' in cell_ref and cell_ref.startswith('$')
        row_abs = '$' in cell_ref and cell_ref.rfind('$') != 0

        # 去掉 $ 后再处理
        ref = cell_ref.replace('$', '')

        # 解析列和行
        col_part = re.match(r'[A-Z]+', ref).group()
        row_part = re.search(r'\d+', ref).group()

        # 转换列字母为数字
        col_num = 0
        for char in col_part:
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)

        # # 如果不是绝对列，则偏移
        # if not col_abs:
        #     col_num += col_shift
        #
        # # 如果不是绝对行，则偏移
        # if not row_abs:
        #     new_row_num = int(row_part) + row_shift
        # else:
        #     new_row_num = int(row_part)

        col_num += col_shift
        new_row_num = int(row_part) + row_shift

        # 将列数字转换回字母
        new_col_ref = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            new_col_ref = chr(65 + remainder) + new_col_ref

        # 加回绝对引用符号
        col_prefix = '$' if col_abs else ''
        row_prefix = '$' if row_abs else ''

        return f"{col_prefix}{new_col_ref}{row_prefix}{new_row_num}"

    def generate_formulas(self, operations):
        """在同一个工作簿中生成Excel公式"""
        # 确保所有工作表都存在
        for op in operations:
            target_sheet_name = op['target']['sheet']
            if target_sheet_name not in self.wb.sheetnames:
                print(f"创建新工作表: {target_sheet_name}")
                self.wb.create_sheet(title=target_sheet_name)

        # 执行所有公式生成操作
        for op in operations:
            # 获取目标工作表
            target_sheet = self.wb[op['target']['sheet']]

            # 处理循环操作
            if 'loop' in op:
                loop_count = op['loop']['count']
                for i in range(loop_count):
                    # 准备公式参数
                    param_refs = {}
                    for param_name, param_def in op['params'].items():
                        # 获取该参数的偏移设置
                        param_offset = op['loop'].get('param_offsets', {}).get(param_name, {})
                        col_shift = param_offset.get('col_shift', 0) * i
                        row_shift = param_offset.get('row_shift', 0) * i

                        # 应用偏移
                        shifted_cell = self.shift_cell(
                            param_def['cell'],
                            col_shift=col_shift,
                            row_shift=row_shift
                        )
                        # 生成参数引用
                        param_refs[param_name] = self.get_formula_reference(
                            param_def['sheet'],
                            shifted_cell
                        )

                    # 获取目标单元格的偏移设置
                    target_offset = op['loop'].get('target_offset', {})
                    target_col_shift = target_offset.get('col_shift', 0) * i
                    target_row_shift = target_offset.get('row_shift', 0) * i

                    # 计算目标单元格位置
                    target_cell = self.shift_cell(
                        op['target']['cell'],
                        col_shift=target_col_shift,
                        row_shift=target_row_shift
                    )

                    # 生成Excel公式（支持内置函数）
                    formula = self.generate_excel_formula(op, param_refs)

                    # 写入公式
                    target_sheet[target_cell] = formula
                    print(f"写入公式: {target_sheet.title}[{target_cell}] {formula}")
            else:
                # 非循环操作处理
                target_cell = op['target']['cell']

                # 准备公式参数
                param_refs = {}
                for param_name, param_def in op['params'].items():
                    param_refs[param_name] = self.get_formula_reference(
                        param_def['sheet'],
                        param_def['cell']
                    )

                # 生成Excel公式（支持内置函数）
                formula = self.generate_excel_formula(op, param_refs)

                # 写入公式
                target_sheet[target_cell] = formula
                print(f"写入公式: {target_sheet.title}[{target_cell}] = {formula}")

        # 保存工作簿
        self.wb.save(self.output_file)
        print(f"公式已生成并保存到: {self.output_file}")
        return self.output_file
    def generate_excel_formula(self, operation, param_refs):
        """根据操作类型生成Excel公式，支持内置函数"""
        op_type = operation['operation']

        if op_type == 'excel_sum':
            # SUM函数: =SUM(range)
            range_ref = param_refs['range']
            return f"=SUM({range_ref})"

        elif op_type == 'excel_average':
            # AVERAGE函数: =AVERAGE(range)
            range_ref = param_refs['range']
            return f"=AVERAGE({range_ref})"

        elif op_type == 'excel_vlookup':
            # VLOOKUP函数: =VLOOKUP(lookup_value, table_array, col_index, [range_lookup])
            lookup_value = param_refs['lookup_value']
            table_array = param_refs['table_array']
            col_index = operation.get('col_index', 2)
            range_lookup = operation.get('range_lookup', 'FALSE')
            return f"=VLOOKUP({lookup_value}, {table_array}, {col_index}, {range_lookup})"

        elif op_type == 'excel_if':
            # IF函数: =IF(logical_test, value_if_true, value_if_false)
            logical_test = operation['logical_test'].format(**param_refs)
            value_if_true = operation['value_if_true'].format(**param_refs)
            value_if_false = operation['value_if_false'].format(**param_refs)
            return f"=IF({logical_test}, {value_if_true}, {value_if_false})"

        elif op_type == 'excel_sumif':
            # SUMIF函数: =SUMIF(range, criteria, [sum_range])
            range_ref = param_refs['range']
            criteria = operation['criteria'].format(**param_refs)
            sum_range = param_refs.get('sum_range', '')
            if sum_range:
                return f"=SUMIF({range_ref}, {criteria}, {sum_range})"
            else:
                return f"=SUMIF({range_ref}, {criteria})"

        elif op_type == 'mixed_offset_calc':
            # 自定义计算: =(A * B) + (C * D)
            return f"=({param_refs['A_val']} * {param_refs['B_val']}) + ({param_refs['C_val']} * {param_refs['D_val']})"

        elif op_type == 'custom':
            # 完全自定义公式模板
            return operation['formula_template'].format(**param_refs)

        else:
            raise ValueError(f"未知的操作类型: {op_type}")


# 使用示例 =============================================
if __name__ == "__main__":
    # 初始化公式生成器（输入输出为同一个文件）
    formula_generator = ExcelFormulaGenerator(
        data_file="多表循环计算.xlsx",  # 包含所有数据的单一文件
        output_file="多表循环计算.xlsx"  # 输出到同一个文件
    )


#  ********************* custom 表示自定义公式内容，如果使用固定好的公式或者excel中的公式计算需要在上方添加 ****************************
#  row_shift 向下移动， col_shift 向右移动

    # 定义公式生成操作 - 独立偏移方向
    operations = [
        # 操作: (A表的x * B表的y) + (C表的z * D表的w)
        # 表A每次向下移动一行 (row_shift=1)
        # 表B、C、D每次向右移动一列 (col_shift=1)
        # 目标单元格每次每次向右移动一列且向下移动一行 (row_shift=1, col_shift=1)
        {
            # "operation": "mixed_offset_calc",  # 混合偏移计算
            "operation": "custom",
            # "formula_template": "=(({A_val} * {B_val}) + ({C_val} * {D_val}))",
            "formula_template": "=({B_val} + {C_val} + {D_val})",
            "params": {
                # "A_val": {"sheet": "Sheet5", "cell": "F5"},  # A表的x (向下移动)
                "B_val": {"sheet": "Sheet5", "cell": "F5"},  # B表的y (向右移动)
                "C_val": {"sheet": "Sheet5", "cell": "G5"},  # C表的z (向右移动)
                "D_val": {"sheet": "Sheet5", "cell": "H5"}  # D表的w (向右移动)
            },
            "target": {
                "sheet": "Sheet5",  # 结果工作表
                "cell": "F10"  # 目标起始位置 (向右移动)
            },
            "loop": {
                "count": 3,  # 循环3次

                # 参数独立偏移设置
                "param_offsets": {
                    # "A_val": {"row_shift": 1, "col_shift": 0},  # 每次向下移动一行
                    "B_val": {"row_shift": 1, "col_shift": 0},  # 每次向右移动一列
                    "C_val": {"row_shift": 1, "col_shift": 0},  # 每次向右移动一列
                    "D_val": {"row_shift": 1, "col_shift": 0}  # 每次向右移动一列
                },

                # 目标单元格偏移设置
                "target_offset": {"row_shift": 0, "col_shift": 1}  # 每次向右移动一列且向下移动一行
            }
        }

        , {
            # "operation": "mixed_offset_calc",  # 混合偏移计算
            "operation": "custom",
            # "formula_template": "=(({A_val} * {B_val}) + ({C_val} * {D_val}))",
            "formula_template": "=({B_val} + {C_val} + SUM({D_val}))",
            "params": {
                # "A_val": {"sheet": "Sheet5", "cell": "F5"},  # A表的x (向下移动)
                "B_val": {"sheet": "Sheet5", "cell": "F5"},  # B表的y (向右移动)
                "C_val": {"sheet": "Sheet5", "cell": "G5"},  # C表的z (向右移动)
                "D_val": {"sheet": "Sheet4", "cell": "E4:G4"}  # D表的w (向右移动)
            },
            "target": {
                "sheet": "Sheet5",  # 结果工作表
                "cell": "F12"  # 目标起始位置 (向右移动)
            },
            "loop": {
                "count": 3,  # 循环3次

                # 参数独立偏移设置
                "param_offsets": {
                    # "A_val": {"row_shift": 1, "col_shift": 0},  # 每次向下移动一行
                    "B_val": {"row_shift": 1, "col_shift": 0},  # 每次向右移动一列
                    "C_val": {"row_shift": 1, "col_shift": 0},  # 每次向右移动一列
                    "D_val": {"row_shift": 1, "col_shift": 0}  # 每次向右移动一列
                },

                # 目标单元格偏移设置
                "target_offset": {"row_shift": 0, "col_shift": 1}  # 每次向右移动一列且向下移动一行
            }
        }

        , {
            "operation": "excel_sum",
            "params": {
                 "range": {"sheet": "Sheet1", "cell": "F3:F7"}  # 区域引用
                },
            "target": {
                "sheet": "Sheet5",
                "cell": "F15"
            },
            "loop": {
                "count": 3,  # 循环3次
                # 参数独立偏移设置
                "param_offsets": {
                    "range": {"row_shift": 1, "col_shift": 0},  # 每次向下移动一行
                },
                # 目标单元格偏移设置
                "target": {"row_shift": 0, "col_shift": 1}  # 每次向右移动一列且向下移动一行
            }
        }

        , {
            # "operation": "mixed_offset_calc",  # 混合偏移计算
            "operation": "custom",
            # "formula_template": "=(({A_val} * {B_val}) + ({C_val} * {D_val}))",
            "formula_template": "=MAX({B_val},{C_val})",
            "params": {
                # "A_val": {"sheet": "Sheet5", "cell": "F5"},  # A表的x (向下移动)
                "B_val": {"sheet": "Sheet5", "cell": "F5"},  # B表的y (向右移动)
                "C_val": {"sheet": "Sheet5", "cell": "G5"},  # C表的z (向右移动)
            },
            "target": {
                "sheet": "Sheet5",  # 结果工作表
                "cell": "F13"  # 目标起始位置 (向右移动)
            },
            "loop": {
                "count": 3,  # 循环3次

                # 参数独立偏移设置
                "param_offsets": {
                    # "A_val": {"row_shift": 1, "col_shift": 0},  # 每次向下移动一行
                    "B_val": {"row_shift": 1, "col_shift": 0},  # 每次向右移动一列
                    "C_val": {"row_shift": 1, "col_shift": 0},  # 每次向右移动一列
                },

                # 目标单元格偏移设置
                "target_offset": {"row_shift": 0, "col_shift": 1}  # 每次向右移动一列且向下移动一行
            }
        }
        #
        # # 操作2: 自定义公式（A表向下移动，B表向右移动）
        # , {
        #     "operation": "custom",
        #     "formula_template": "=({A_val} * {B_val}) + ({C_val} / {D_val})",
        #     "params": {
        #         "A_val": {"sheet": "Sheet1", "cell": "B2"},  # 每次向下移动
        #         "B_val": {"sheet": "Sheet2", "cell": "C3"},  # 每次向右移动
        #         "C_val": {"sheet": "Sheet3", "cell": "D4"},  # 每次向右移动
        #         "D_val": {"sheet": "Sheet4", "cell": "E5"}  # 每次向右移动
        #     },
        #     "target": {
        #         "sheet": "结果表",
        #         "cell": "G5"  # 目标起始位置
        #     },
        #     "loop": {
        #         "count": 5,
        #         "param_offsets": {
        #             "A_val": {"row_shift": 1, "col_shift": 0},  # 每次向下移动一行
        #             "B_val": {"row_shift": 0, "col_shift": 1},  # 每次向右移动一列
        #             "C_val": {"row_shift": 0, "col_shift": 1},  # 每次向右移动一列
        #             "D_val": {"row_shift": 0, "col_shift": 1}  # 每次向右移动一列
        #         },
        #         "target_offset": {"row_shift": 0, "col_shift": 1}  # 每次向右移动一列
        #     }
        # }

        # , {
        #     "operation": "excel_if",
        #     "params": {
        #         "sales": {"sheet": "销售表", "cell": "B2"},
        #         "target": {"sheet": "目标表", "cell": "C2"}
        #     },
        #     "logical_test": "{sales} > {target}",
        #     "value_if_true": "\"达标\"",
        #     "value_if_false": "\"未达标\"",
        #     "target": {
        #         "sheet": "结果表",
        #         "cell": "D2"
        #     }
        # }
    ]

    # 生成公式
    formula_generator.generate_formulas(operations)