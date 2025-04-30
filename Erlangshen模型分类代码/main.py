import pandas as pd
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import torch
import os


def analyze_sentiment_batch(texts, tokenizer, model, labels):
    """使用模型进行批量情感分析"""
    # 确保texts是字符串列表
    if not isinstance(texts, list) or not all(isinstance(text, str) for text in texts):
        raise ValueError("texts must be a list of strings")
    inputs = tokenizer(texts, return_tensors="pt", truncation=True, padding=True, max_length=512)
    outputs = model(**inputs)
    probs = torch.nn.functional.softmax(outputs.logits, dim=-1)
    predicted_labels = torch.argmax(probs, dim=1).tolist()
    confidences = [probs[i, predicted_label].item() for i, predicted_label in enumerate(predicted_labels)]
    # 返回标签索引和置信度
    return [(predicted_label, confidence) for predicted_label, confidence in zip(predicted_labels, confidences)]


def main():
    # 本地模型路径
    LOCAL_MODEL_PATH = r"E:\NLP_Finetune\training_output_20250128_024658\model"  # 替换为你的本地路径

    # 检查路径是否存在
    if not os.path.exists(LOCAL_MODEL_PATH):
        print(f"路径 {LOCAL_MODEL_PATH} 不存在，请检查！")
        return

    # 加载模型和分词器
    try:
        tokenizer = AutoTokenizer.from_pretrained(LOCAL_MODEL_PATH, local_files_only=True)
        model = AutoModelForSequenceClassification.from_pretrained(LOCAL_MODEL_PATH, local_files_only=True)
    except Exception as e:
        print(f"加载本地模型或分词器时出错: {e}")
        return

    # 定义情感标签
    labels = ["负面" , "正面"]

    # 输入输出文件路径
    input_file = r'D:\桌面\爬取数据\预处理\名创优品\名创优品补抓数据25.2.13预处理_326条.xlsx'
    output_file = r'D:\桌面\爬取数据\预处理\名创优品\名创优品补抓数据25.2.13情感分析结果.xlsx'

    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"读取文件 {input_file} 时出错: {e}")
        return

    if "text" not in df.columns:
        print("Excel 文件中没有 'text' 列。请检查文件格式。")
        return

    # 过滤空值和非字符串
    df["text"] = df["text"].fillna("").astype(str)

    # 对文本列进行情感分析
    batch_size = 16  # 可根据内存情况调整
    results = []
    for i in range(0, len(df["text"]), batch_size):
        batch_texts = df["text"][i:i + batch_size].tolist()
        # 过滤掉空字符串的批次
        batch_texts = [text for text in batch_texts if text.strip()]
        if not batch_texts:  # 如果批次为空，跳过
            print(f"跳过空批次 {i}-{i + batch_size}")
            continue
        try:
            batch_results = analyze_sentiment_batch(batch_texts, tokenizer, model, labels)
            results.extend(batch_results)
        except Exception as e:
            print(f"处理批次 {i}-{i + batch_size} 时出错: {e}")
            continue

    # 将分析结果添加到 DataFrame
    results_df = pd.DataFrame(results, columns=["LabelIndex", "Confidence"])
    df = pd.concat([df, results_df], axis=1)

    # 保存结果到新文件
    try:
        with pd.ExcelWriter(output_file) as writer:
            df.to_excel(writer, index=False)
        print(f"分析结果已保存到 {output_file}")
    except Exception as e:
        print(f"保存文件 {output_file} 时出错: {e}")


if __name__ == "__main__":
    main()