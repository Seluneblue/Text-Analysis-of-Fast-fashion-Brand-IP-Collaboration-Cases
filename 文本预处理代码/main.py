import re
import pandas as pd
from bs4 import BeautifulSoup
from pathlib import Path
import logging
import emoji

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 去除 HTML 标签
def remove_html(text):
    if isinstance(text, (str, bytes, Path)) and Path(text).is_file():  # 检查输入是否为文件路径，且类型正确
        with open(text, 'r', encoding='utf-8') as file:
            text = file.read()
    return BeautifulSoup(text, "html.parser").get_text()

# 去除 URL
def remove_urls(text):
    return re.sub(r'http\S+', '', text)

# 去除表情符号，使用 emoji 库
def remove_emojis(text):
    return emoji.replace_emoji(text, replace='')

# 去除特殊字符，添加对 {%...%} 的处理
def remove_special_characters(text, allowed_chars='\w\s,\u4e00-\u9fa5@'):
    # 调整正则表达式，更精确地处理 {%...%} 结构
    pattern = re.compile(r'[^{}\w\s,\u4e00-\u9fa5@]+|{%[^}]+?%}')
    return pattern.sub('', text)

# 新增：去除 @xxx 直到空格或常见标点符号的内容
def remove_at_mentions(text):
    logging.info(f"Before remove_at_mentions: {text}")
    result = re.sub(r'@.*?(?=\s|$)', '', text)
    logging.info(f"After remove_at_mentions: {result}")
    return result

# 读取 Excel 文件并处理文本
def preprocess_text_in_excel(input_path, output_path, text_column='text'):
    try:
        # 使用 pathlib 处理文件路径
        input_path = Path(input_path)
        output_path = Path(output_path)
        # 读取 Excel 文件
        df = pd.read_excel(input_path, engine='openpyxl')
        # 检查是否包含文本列
        if text_column not in df.columns:
            raise KeyError(f"错误：Excel 中没有 '{text_column}' 列")
        # 对文本列进行处理
        df[text_column] = df[text_column].apply(preprocess_text)
        # 保存处理后的 DataFrame 到新的 Excel 文件
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False)
        logging.info(f"预处理完成，处理后的文件已保存到: {output_path}")
    except Exception as e:
        logging.error(f"处理文件时发生错误: {str(e)}")

# 文本清理
def preprocess_text(text):
    if isinstance(text, (float, int)):  # 检查是否为 float 或 int 类型
        text = str(text)  # 将 float 或 int 类型转换为字符串
    logging.info(f"原始文本: {text}")  # 记录原始文本
    text = remove_html(text)  # 去除 HTML 标签
    logging.info(f"去除 HTML 标签后: {text}")  # 记录去除 HTML 标签后的文本
    text = remove_urls(text)  # 去除 URL
    logging.info(f"去除 URL 后: {text}")  # 记录去除 URL 后的文本
    text = remove_emojis(text)  # 去除表情符号
    logging.info(f"去除表情符号后: {text}")  # 记录去除表情符号后的文本
    text = remove_special_characters(text)  # 去除特殊字符
    logging.info(f"去除特殊字符后: {text}")  # 记录去除特殊字符后的文本
    # 新增：去除 @xxx 直到空格或常见标点符号的内容
    text = remove_at_mentions(text)
    logging.info(f"去除 @xxx 直到空格或常见标点符号的内容后: {text}")  # 记录去除 @xxx 后的文本
    return text

# 文件路径
input_path = r"D:\桌面\爬取数据\预处理\名创优品\名创优品补抓数据25.2.13_utf8.xlsx"  # 输入文件路径
output_path = r"D:\桌面\爬取数据\预处理\名创优品\名创优品补抓数据25.2.13预处理.xlsx"  # 输出文件路径

# 调用函数进行预处理
preprocess_text_in_excel(input_path, output_path)