from loan_calculate import loan_fill
from copy_formula import just_copy, special_copy, clear_style_cache
from other_table import copy_one, copy_two, copy_three
from final_table import final_copy, table_c_last
from openpyxl_vba import load_workbook
import shutil
import os
import win32com.client as win32


def restore_controls_and_macros(file_path, source_file):
    """使用win32com恢复控件和宏"""
    # 启动Excel应用程序
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:

        # 打开备份文件
        wb_backup = excel.Workbooks.Open(os.path.abspath(source_file))

        # 打开处理后的文件
        wb_processed = excel.Workbooks.Open(os.path.abspath(file_path))

        sheet_name = "A财务假设"
        sheet_backup = wb_backup.Sheets[sheet_name]

        # 获取备份工作表中的所有控件
        try:
            # 尝试获取表单控件
            for control in sheet_backup.Shapes:
                # 复制控件到处理后的工作表
                control.Copy()
                wb_processed.Sheets(sheet_name).Paste()
        except Exception as e:
            print(f"复制控件时出错: {e}")

        # 保存并关闭工作簿
        wb_processed.Save()
        wb_processed.Close()
        wb_backup.Close()

        print("控件和宏已恢复")

    except Exception as e:
        print(f"恢复控件和宏时出错: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # 确保关闭Excel应用程序
        excel.Quit()



if __name__ == "__main__":
    source_path = "财务分析套表自做模板编程用ver5.xlsm"
    input_path = "财务分析套表自做模板编程用2.xlsm"
    # last_year = 2047

    # 复制文件（保留所有元数据和宏）
    if os.path.exists(input_path):
        os.remove(input_path)

    shutil.copy2(source_path, input_path)

    last_year = loan_fill(input_path)
    copy_three(input_path, last_year)
    copy_two(input_path, last_year)
    copy_one(input_path, last_year)
    just_copy(input_path, last_year)
    special_copy(input_path, last_year)
    final_copy(input_path, last_year)
    table_c_last(input_path)
    clear_style_cache()

    # if os.path.exists(input_path) and os.path.exists(source_path):
    #     print("恢复控件和宏...")
    #     restore_controls_and_macros(input_path, source_path)
    print("done")

