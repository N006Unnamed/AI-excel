import tkinter as tk
from tkinter import filedialog, ttk
import time
import threading
from datetime import timedelta
from loan_calculate import loan_fill
from copy_formula import just_copy, special_copy, clear_style_cache
from other_table import copy_one, copy_two, copy_three
from final_table import final_copy, table_c_last
import shutil
import os


# 假设这是您已经完成的Excel修改功能
# 请确保您已经定义了以下函数：
# loan_fill, copy_three, copy_two, copy_one, just_copy, special_copy, final_copy, table_c_last, clear_style_cache

def modify_excel_file(input_path, output_path, progress_callback=None):
    """
    修改Excel文件的主函数
    :param input_path: 输入文件路径
    :param output_path: 输出文件路径
    :param progress_callback: 进度回调函数，用于更新GUI进度
    """
    try:
        if progress_callback:
            progress_callback(1, "正在读取文件")

        # 复制文件（保留所有元数据和宏）
        if os.path.exists(output_path):
            os.remove(output_path)

        shutil.copy2(input_path, output_path)

        if progress_callback:
            progress_callback(12.5, "正在生成表c...")
        last_year = loan_fill(output_path)

        if progress_callback:
            progress_callback(25, "正在生成表E.2、E.3...")
        copy_three(output_path, last_year)

        if progress_callback:
            progress_callback(37.5, "正在生成表E.1、E.1.1...")
        copy_two(output_path, last_year)

        if progress_callback:
            progress_callback(50, "正在生成表D.2、D.4、E、F、F.1、G...")
        copy_one(output_path, last_year)

        if progress_callback:
            progress_callback(62.5, "正在生成表b、d、e...")
        just_copy(output_path, last_year)

        if progress_callback:
            progress_callback(75, "正在生成表a.1、a.2...")
        special_copy(output_path, last_year)

        if progress_callback:
            progress_callback(87.5, "正在生成表III、VI...")
        final_copy(output_path, last_year)

        if progress_callback:
            progress_callback(99, "正在回填表c...")
        table_c_last(output_path)

        if progress_callback:
            progress_callback(100, "完成")
        clear_style_cache()

        print("处理完成")
        return True, f"处理完成，文件已保存至: {output_path}"

    except Exception as e:
        # 如果出错，删除可能已创建的部分输出文件
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass
        return False, f"处理过程中发生错误: {str(e)}"


class ExcelModifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("财务报表智能生成")
        self.root.geometry("600x400")  # 增加宽度和高度以容纳新控件

        # 变量初始化
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.output_file_name = tk.StringVar(value="output.xlsm")
        self.is_processing = False
        self.start_time = None
        self.timer_running = False

        self.create_widgets()

    def create_widgets(self):
        # 文件选择部分
        file_frame = tk.LabelFrame(self.root, text="文件设置", padx=10, pady=10)
        file_frame.pack(padx=10, pady=5, fill="x")

        # 输入文件选择
        tk.Label(file_frame, text="输入 Excel 文件:").grid(row=0, column=0, sticky="w", pady=5)
        input_frame = tk.Frame(file_frame)
        input_frame.grid(row=0, column=1, sticky="ew", pady=5)
        file_frame.columnconfigure(1, weight=1)

        tk.Entry(input_frame, textvariable=self.input_file_path).pack(side="left", fill="x", expand=True)
        tk.Button(input_frame, text="浏览", command=self.browse_input_file, width=8).pack(side="right", padx=(5, 0))

        # 输出文件名设置
        tk.Label(file_frame, text="输出文件名:").grid(row=1, column=0, sticky="w", pady=5)
        name_frame = tk.Frame(file_frame)
        name_frame.grid(row=1, column=1, sticky="ew", pady=5)

        tk.Entry(name_frame, textvariable=self.output_file_name).pack(side="left", fill="x", expand=True)

        # 输出路径选择
        tk.Label(file_frame, text="保存路径:").grid(row=2, column=0, sticky="w", pady=5)
        output_frame = tk.Frame(file_frame)
        output_frame.grid(row=2, column=1, sticky="ew", pady=5)

        tk.Entry(output_frame, textvariable=self.output_file_path).pack(side="left", fill="x", expand=True)
        tk.Button(output_frame, text="浏览", command=self.browse_output_path, width=8).pack(side="right", padx=(5, 0))

        # 按钮区域
        button_frame = tk.Frame(self.root, padx=10, pady=10)
        button_frame.pack(padx=10, pady=5)

        self.process_button = tk.Button(
            button_frame,
            text="开始运行程序",
            command=self.start_processing,
            width=15,
            height=2,
            bg="#4CAF50",
            fg="white",
            font=("", 10, "bold")
        )
        self.process_button.pack()

        # 计时器区域
        timer_frame = tk.Frame(self.root, padx=10, pady=5)
        timer_frame.pack(padx=10, pady=5, fill="x")

        tk.Label(timer_frame, text="运行时间:").pack(anchor="w")
        self.timer_label = tk.Label(timer_frame, text="00:00:00", font=("Arial", 10, "bold"))
        self.timer_label.pack(anchor="w")

        # 进度条区域
        progress_frame = tk.LabelFrame(self.root, text="处理进度", padx=10, pady=10)
        progress_frame.pack(padx=10, pady=5, fill="x")

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate")
        self.progress_bar.pack(fill="x", pady=5)

        self.status_label = tk.Label(progress_frame, text="就绪")
        self.status_label.pack(anchor="w")

    def browse_input_file(self):
        if self.is_processing:
            return

        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
        )
        if filename:
            self.input_file_path.set(filename)
            self.update_status(0, "就绪")

            # 自动设置默认输出文件名
            if not self.output_file_name.get() or self.output_file_name.get() == "处理后的文件.xlsm":
                base_name = os.path.splitext(os.path.basename(filename))[0]
                self.output_file_name.set(f"{base_name}_处理结果.xlsm")

            # 自动设置默认输出路径为输入文件所在目录
            if not self.output_file_path.get():
                self.output_file_path.set(os.path.dirname(filename))

    def browse_output_path(self):
        if self.is_processing:
            return

        path = filedialog.askdirectory(
            title="选择保存目录"
        )
        if path:
            self.output_file_path.set(path)

    def update_status(self, progress, message):
        """更新进度和状态信息"""
        # 确保进度值在0-100之间
        progress = max(0, min(100, progress))
        self.status_label.config(text=message)
        self.progress_bar["value"] = progress
        self.root.update_idletasks()

    def update_timer(self):
        """更新计时器显示"""
        if self.timer_running:
            elapsed = time.time() - self.start_time
            # 将秒转换为时:分:秒格式
            elapsed_str = str(timedelta(seconds=int(elapsed)))
            self.timer_label.config(text=elapsed_str)
            # 每秒更新一次
            self.root.after(1000, self.update_timer)

    def start_timer(self):
        """启动计时器"""
        self.start_time = time.time()
        self.timer_running = True
        self.update_timer()

    def stop_timer(self):
        """停止计时器"""
        self.timer_running = False

    def reset_timer(self):
        """重置计时器"""
        self.timer_label.config(text="00:00:00")

    def start_processing(self):
        """开始处理Excel文件"""
        if self.is_processing:
            return

        if not self.input_file_path.get():
            self.update_status(0, "错误: 请先选择Excel文件")
            return

        if not os.path.isfile(self.input_file_path.get()):
            self.update_status(0, "错误: 输入文件不存在")
            return

        if not self.output_file_path.get():
            self.update_status(0, "错误: 请选择保存路径")
            return

        if not self.output_file_name.get():
            self.update_status(0, "错误: 请输入输出文件名")
            return

        # 构建完整的输出文件路径
        output_path = os.path.join(self.output_file_path.get(), self.output_file_name.get())

        # 检查输出文件是否已存在
        if os.path.exists(output_path):
            # 可以在这里添加确认覆盖的对话框
            pass

        # 重置并启动计时器
        self.reset_timer()
        self.start_timer()

        # 禁用按钮，防止重复点击
        self.is_processing = True
        self.process_button.config(state="disabled", bg="#cccccc")

        # 在新线程中处理Excel文件，避免GUI冻结
        thread = threading.Thread(target=self.process_excel, args=(output_path,))
        thread.daemon = True
        thread.start()

    def process_excel(self, output_path):
        """处理Excel文件的方法"""
        try:
            # 调用您的Excel修改函数
            success, message = modify_excel_file(
                self.input_file_path.get(),
                output_path,
                progress_callback=self.update_status
            )

            if success:
                self.update_status(100, message)
            else:
                self.update_status(0, message)

        except Exception as e:
            self.update_status(0, f"处理过程中发生错误: {str(e)}")

        finally:
            # 停止计时器
            self.stop_timer()

            # 重新启用按钮
            self.is_processing = False
            self.root.after(0, lambda: self.process_button.config(state="normal", bg="#4CAF50"))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelModifierApp(root)
    root.mainloop()