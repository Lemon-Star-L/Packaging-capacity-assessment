# @Author: YiLong Li
# @Date: 2026/10/23 11:10
# @File: clean.py
# @Software: PyCharm


"""
主要功能：
    读取基础资料中的件号信息，清理空白及重复件号，
    形成唯一件号清单。

处理规则：
    1. 清理列名及件号首尾空白。
    2. 删除件号为空的整行。
    3. 按件号去重，保留首次出现的完整记录。
    4. 保留其他字段原始内容，不合并信息、不汇总数量。
    5. 按文本格式读取，保留件号前导零。

输入位置：
    clean/模块信息.xlsx，默认读取第一个工作表。

输出位置：
    clean/模块信息_清洗.xlsx。

使用说明：
    控制台显示原始行数、空白件号删除数、重复件号删除数
    及最终保留行数。
    不修改原始文件，同名旧清洗结果会被覆盖。
    适用于一件号一条记录的基础主数据整理，
    不适用于直接清理业务交易流水。
"""

from pathlib import Path
import pandas as pd

# 读取和输出均使用脚本所在文件夹下的 clean 子文件夹
base_dir = Path(__file__).resolve().parent / "clean"
base_dir.mkdir(parents=True, exist_ok=True)

input_file = base_dir / "模块信息.xlsx"
output_file = base_dir / "模块信息_清洗.xlsx"

sheet_name = 0
column_name = "件号"

if not input_file.is_file():
    raise FileNotFoundError(f"请确认文件存在：{input_file}")

# 按文本读取，保留文本件号中的前导零
df = pd.read_excel(
    input_file,
    sheet_name=sheet_name,
    dtype=str,
    keep_default_na=False
)

# 清理列名首尾空格
df.columns = df.columns.str.strip()

if column_name not in df.columns:
    raise ValueError(
        f"未找到“{column_name}”字段，当前字段为：{df.columns.tolist()}"
    )

original_count = len(df)

# 清理件号首尾空白字符
df[column_name] = df[column_name].str.strip()

# 1. 删除件号为空的整行
empty_mask = df[column_name].eq("")
empty_count = int(empty_mask.sum())
df = df.loc[~empty_mask].copy()

# 2. 按件号去重，保留第一次出现的整行
duplicate_count = int(df.duplicated(subset=[column_name]).sum())
df = df.drop_duplicates(subset=[column_name], keep="first")

# 保存结果，不覆盖原文件
df.to_excel(output_file, index=False)

print(f"原始数据：{original_count} 行")
print(f"删除空白件号：{empty_count} 行")
print(f"删除重复件号：{duplicate_count} 行")
print(f"最终保留：{len(df)} 行")
print(f"文件已保存：{output_file}")