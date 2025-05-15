# Excel 数据处理工具套件

## 项目概述

这个项目包含两个独立的 Excel 数据处理工具，使用 Python 和 Tkinter 开发，具有图形用户界面：

1. **Excel 列合并工具** - 将多个 Excel 文件中指定的列合并为一个新列
2. **Excel 数据处理工具** - 从合并后的数据中提取统计信息（正确/错误人数）

## 功能特点

### Excel 列合并工具
- 选择文件夹批量处理多个 Excel 文件
- 可视化选择需要合并的列
- 自定义合并后的列名
- 自动检测 Excel 文件中的列名
- 友好的用户界面和状态提示

### Excel 数据处理工具
- 选择单个 Excel 文件进行处理
- 从合并后的数据中提取正确/错误人数统计
- 自动生成详细数据和汇总表
- 实时处理进度显示
- 完整的处理日志记录
- 结果自动保存为新的 Excel 文件

## 安装与使用

### 系统要求
- Python 3.6+
- 以下 Python 库：
  - pandas
  - openpyxl
  - tkinter

### 安装依赖
```bash
pip install pandas openpyxl
```

### 使用方法
1. 运行 `ExcelMergerApp.py` 启动列合并工具
   ```bash
   python ExcelMergerApp.py
   ```
2. 运行 `ExcelProcessorApp.py` 启动数据处理工具
   ```bash
   python ExcelProcessorApp.py
   ```

## 界面截图
![c0435a3fc6bfeeb41ee31674d4382bc](https://github.com/user-attachments/assets/7e169c2f-03f0-4cca-92e7-802d41aa4d1a)


## 代码结构

```
excel-tools/
├── ExcelMergerApp.py        # 列合并工具主程序
├── ExcelProcessorApp.py     # 数据处理工具主程序
├── README.md                # 项目文档
└── screenshots/             # 界面截图目录
```

## 开发说明

### 技术栈
- Python 3
- pandas (数据处理)
- Tkinter (GUI)
- openpyxl (Excel 文件操作)

### 自定义修改
- 修改 `ExcelMergerApp.py` 中的 `new_column_name` 默认值可以更改合并列的默认名称
- 修改 `ExcelProcessorApp.py` 中的正则表达式可以适配不同的数据格式

## 贡献指南

欢迎提交 Pull Request 或 Issue 来改进这个项目。主要改进方向包括：
- 增加更多数据处理功能
- 改进用户界面
- 优化性能
- 添加更多文件格式支持

## 联系方式

如有任何问题，请联系项目维护者或提交 Issue。
