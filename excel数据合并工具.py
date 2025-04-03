import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel列合并工具")
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        # 变量
        self.folder_path = tk.StringVar()
        self.merge_columns = []
        self.new_column_name = tk.StringVar(value="合并后的数据")

        # 创建界面
        self.create_widgets()

    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="Excel列合并工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # 文件夹选择
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(pady=10, fill="x", padx=20)

        folder_label = tk.Label(folder_frame, text="选择文件夹:", font=("Arial", 12))
        folder_label.pack(side=tk.LEFT, padx=5)

        folder_entry = tk.Entry(folder_frame, textvariable=self.folder_path, width=40)
        folder_entry.pack(side=tk.LEFT, padx=5)

        folder_button = tk.Button(folder_frame, text="浏览", command=self.select_folder)
        folder_button.pack(side=tk.LEFT, padx=5)

        # 列选择
        column_frame = tk.Frame(self.root)
        column_frame.pack(pady=10, fill="x", padx=20)

        column_label = tk.Label(column_frame, text="选择要合并的列:", font=("Arial", 12))
        column_label.pack(side=tk.LEFT, padx=5)

        self.column_listbox = tk.Listbox(column_frame, selectmode=tk.MULTIPLE, width=30, height=10)
        self.column_listbox.pack(side=tk.LEFT, padx=5, fill="both", expand=True)

        # 新列名
        new_column_frame = tk.Frame(self.root)
        new_column_frame.pack(pady=10, fill="x", padx=20)

        new_column_label = tk.Label(new_column_frame, text="合并后列名:", font=("Arial", 12))
        new_column_label.pack(side=tk.LEFT, padx=5)

        new_column_entry = tk.Entry(new_column_frame, textvariable=self.new_column_name, width=30)
        new_column_entry.pack(side=tk.LEFT, padx=5)

        # 操作按钮
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20, fill="x", padx=20)

        process_button = tk.Button(button_frame, text="开始处理", command=self.process_files, bg="#4CAF50", fg="white",
                                   font=("Arial", 12, "bold"), width=15)
        process_button.pack(side=tk.LEFT, padx=10)

        # 状态显示
        self.status_label = tk.Label(self.root, text="请选择文件夹并选择要合并的列", font=("Arial", 10))
        self.status_label.pack(pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self.load_columns(folder)

    def load_columns(self, folder):
        # 尝试从第一个Excel文件中加载列名
        try:
            # 获取文件夹中的所有Excel文件
            excel_files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls', '.csv'))]

            if not excel_files:
                messagebox.showwarning("警告", "文件夹中没有找到Excel文件！")
                return

            # 读取第一个文件的列名
            first_file = os.path.join(folder, excel_files[0])
            df = pd.read_excel(first_file)
            columns = df.columns.tolist()

            # 更新列表框
            self.column_listbox.delete(0, tk.END)
            for column in columns:
                self.column_listbox.insert(tk.END, column)

            self.status_label.config(text=f"已加载列名，共 {len(columns)} 列")

        except Exception as e:
            messagebox.showerror("错误", f"加载列名时出错: {e}")

    def get_selected_columns(self):
        selected_columns = []
        for i in self.column_listbox.curselection():
            selected_columns.append(self.column_listbox.get(i))
        return selected_columns

    def process_files(self):
        folder = self.folder_path.get()
        selected_columns = self.get_selected_columns()
        new_column_name = self.new_column_name.get()

        if not folder:
            messagebox.showwarning("警告", "请选择文件夹！")
            return

        if not selected_columns:
            messagebox.showwarning("警告", "请选择要合并的列！")
            return

        try:
            # 获取文件夹中的所有Excel文件
            excel_files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls', '.csv'))]

            if not excel_files:
                messagebox.showwarning("警告", "文件夹中没有找到Excel文件！")
                return

            # 读取第一个文件作为基准
            first_file = os.path.join(folder, excel_files[0])
            merged_data = pd.read_excel(first_file)
            merged_data[new_column_name] = merged_data[selected_columns].apply(lambda row: ', '.join(row.astype(str)), axis=1)

            # 处理其他文件
            for file in excel_files[1:]:
                file_path = os.path.join(folder, file)
                df = pd.read_excel(file_path)

                # 检查合并列是否存在
                missing_columns = [col for col in selected_columns if col not in df.columns]
                if missing_columns:
                    messagebox.showerror("错误", f"文件 {file} 中没有找到列: {', '.join(missing_columns)}！")
                    return

                # 确保行数一致
                if len(df) != len(merged_data):
                    messagebox.showerror("错误", f"文件 {file} 的行数与第一个文件不一致！")
                    return

                # 合并指定列的数据到新列中
                for idx in range(len(df)):
                    merged_data.loc[idx, new_column_name] += ", " + ', '.join(df.loc[idx, selected_columns].astype(str))

            # 保存结果
            output_file = os.path.join(folder, "merged_result.xlsx")
            merged_data.to_excel(output_file, index=False)

            messagebox.showinfo("完成", f"处理完成！结果已保存到: {output_file}")
            self.status_label.config(text=f"处理完成！结果已保存到: {output_file}")

        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出错: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()