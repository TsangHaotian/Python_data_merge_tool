import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel数据处理工具")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")

        # 创建主框架
        self.main_frame = tk.Frame(root, bg="#f0f0f0")
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 文件选择部分
        self.file_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.file_frame.pack(fill=tk.X, pady=10)

        self.file_label = tk.Label(self.file_frame, text="未选择文件", bg="#f0f0f0", font=("Arial", 12))
        self.file_label.pack(side=tk.LEFT, padx=10)

        self.browse_button = tk.Button(
            self.file_frame,
            text="选择Excel文件",
            command=self.select_file,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5
        )
        self.browse_button.pack(side=tk.RIGHT, padx=10)

        # 处理按钮
        self.process_button = tk.Button(
            self.main_frame,
            text="处理数据",
            command=self.process_data,
            bg="#2196F3",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10,
            state=tk.DISABLED
        )
        self.process_button.pack(pady=20)

        # 进度条
        self.progress_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.progress_frame.pack(fill=tk.X, pady=10)

        self.progress_label = tk.Label(
            self.progress_frame,
            text="处理进度：",
            bg="#f0f0f0",
            font=("Arial", 10)
        )
        self.progress_label.pack(side=tk.LEFT, padx=10)

        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            orient=tk.HORIZONTAL,
            length=500,
            mode='determinate'
        )
        self.progress_bar.pack(side=tk.RIGHT, padx=10)

        # 日志区域
        self.log_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_label = tk.Label(
            self.log_frame,
            text="处理日志",
            bg="#f0f0f0",
            font=("Arial", 12, "bold")
        )
        self.log_label.pack(anchor=tk.W, padx=10, pady=5)

        self.log_text = ScrolledText(
            self.log_frame,
            wrap=tk.WORD,
            font=("Arial", 10),
            bg="white",
            fg="black",
            padx=5,
            pady=5
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10)

        # 初始化变量
        self.file_path = None
        self.correct_total = 0
        self.incorrect_total = 0

    def select_file(self):
        """选择Excel文件"""
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if self.file_path:
            self.file_label.config(text=self.file_path.split('/')[-1])
            self.process_button.config(state=tk.NORMAL)
            self.log(f"已选择文件: {self.file_path}")

    def log(self, message):
        """记录日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.root.update()

    def process_data(self):
        """处理Excel数据"""
        if not self.file_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return

        try:
            # 读取Excel文件
            self.log("开始读取Excel文件...")
            df = pd.read_excel(self.file_path)
            self.log("Excel文件读取完成")

            # 初始化统计变量
            self.correct_total = 0
            self.incorrect_total = 0

            # 正则表达式模式
            pattern = re.compile(r'正确：(\d+)人\s*错误：(\d+)人')

            # 用于存储每行的正确和错误人数
            correct_per_row = []
            incorrect_per_row = []

            # 遍历每一行，提取正确和错误人数
            total_rows = len(df)
            for index, row in df.iterrows():
                merged_data = row.get('合并后的数据', '')
                matches = pattern.findall(merged_data)

                correct = 0
                incorrect = 0

                for match in matches:
                    correct += int(match[0])
                    incorrect += int(match[1])

                correct_per_row.append(correct)
                incorrect_per_row.append(incorrect)

                # 累计总人数
                self.correct_total += correct
                self.incorrect_total += incorrect

                # 更新进度条
                progress = (index + 1) / total_rows * 100
                self.progress_bar['value'] = progress
                self.root.update_idletasks()

            # 创建一个新的DataFrame，包含原始数据和正确、错误人数
            df['正确人数'] = correct_per_row
            df['错误人数'] = incorrect_per_row

            # 汇总结果
            summary = {
                '总正确人数': [self.correct_total],
                '总错误人数': [self.incorrect_total]
            }
            summary_df = pd.DataFrame(summary)

            # 将结果输出到新的Excel文件
            output_file = '修正.xlsx'
            self.log(f"开始将结果保存到 {output_file}...")
            with pd.ExcelWriter(output_file) as writer:
                df.to_excel(writer, sheet_name='详细数据', index=False)
                summary_df.to_excel(writer, sheet_name='汇总', index=False)
            self.log(f"结果已保存到 {output_file}")

            # 显示处理结果
            messagebox.showinfo(
                "处理完成",
                f"处理完成！\n总正确人数：{self.correct_total}\n总错误人数：{self.incorrect_total}\n结果已保存到 {output_file}"
            )

        except Exception as e:
            self.log(f"处理失败: {str(e)}")
            messagebox.showerror("错误", f"处理失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()