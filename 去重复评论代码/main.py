import os
import pandas as pd

# 读取 Excel 文件
file_path = r'D:\桌面\爬取数据\预处理\名创优品\xhs名创优品产品1预处理.xlsx'
df = pd.read_excel(file_path)

# 找出 'dyID' 和 'text' 列都重复的行的索引
duplicate_index = df[df.duplicated(subset=['idname', 'text'])].index

# 删除重复行
df = df.drop(duplicate_index)

# 新文件的保存路径
new_file_path = r'D:\桌面\爬取数据\预处理\名创优品\xhs名创优品产品1预处理_去重复.xlsx'
# 创建新路径的目录
os.makedirs(os.path.dirname(new_file_path), exist_ok=True)

# 将处理后的数据保存到新的 Excel 文件中
df.to_excel(new_file_path, index=False)