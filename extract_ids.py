import pandas as pd
import re
import os

# 读取 Excel 文件（确保文件名与上传的一致）
excel_file = "xhs_hanfu_account_classified.xlsx"
sheet_name = "账号分类数据"   # 你的 Excel 中数据所在的 sheet 名

# 检查文件是否存在
if not os.path.exists(excel_file):
    raise FileNotFoundError(f"找不到文件 {excel_file}，请确认已上传到仓库根目录")

# 读取 Excel
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# 提取笔记 ID（从 "笔记链接" 列）
note_ids = []
for url in df["笔记链接"]:
    if pd.isna(url):
        continue
    match = re.search(r"explore/([a-f0-9]+)", str(url))
    if match:
        note_ids.append(match.group(1))

if not note_ids:
    raise ValueError("未提取到任何笔记 ID，请检查 Excel 中 '笔记链接' 列是否有数据")

print(f"成功提取 {len(note_ids)} 个笔记 ID")

# 读取 config/xhs_config.py 文件
config_path = "config/xhs_config.py"
if not os.path.exists(config_path):
    raise FileNotFoundError(f"找不到配置文件 {config_path}")

with open(config_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

# 找到 XHS_SPECIFIED_ID_LIST 并替换
new_lines = []
inside_list = False
found = False
for line in lines:
    if "XHS_SPECIFIED_ID_LIST" in line and "=" in line:
        found = True
        # 生成新的列表行
        new_lines.append(f'XHS_SPECIFIED_ID_LIST = {note_ids}\n')
        inside_list = True
    elif inside_list and line.strip().startswith("]"):
        inside_list = False
        # 跳过原来的列表内容（因为我们已经用新行替代了）
        continue
    elif inside_list:
        # 跳过原来的列表项
        continue
    else:
        new_lines.append(line)

if not found:
    # 如果文件中没有该变量，则在文件末尾追加
    new_lines.append(f'\n# 自动生成的笔记ID列表\nXHS_SPECIFIED_ID_LIST = {note_ids}\n')

# 写回文件
with open(config_path, "w", encoding="utf-8") as f:
    f.writelines(new_lines)

print("已成功更新 config/xhs_config.py 中的 XHS_SPECIFIED_ID_LIST")
