# Excel 表格合并清洗工具

把多个Excel文件合并成一个，顺便去掉重复数据。

## 能做什么
- 合并多个 xlsx/xls/csv 文件
- 自动去掉重复的行
- 可以按某一列去重（比如按学号）
- 空值自动填成无
- 标记每条数据来自哪个文件

## 怎么用
python excel_merger.py 文件夹路径 输出文件名.xlsx

如果想按某列去重就加上第三个参数:
python excel_merger.py 文件夹路径 输出.xlsx 学号

## 需要的库
pip install pandas openpyxl
