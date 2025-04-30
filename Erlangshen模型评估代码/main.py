import os
import pandas as pd
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import torch
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score


def analyze_sentiment_batch(texts, tokenizer, model):
    """使用模型进行批量情感分析"""
    # 确保texts是字符串列表
    if not isinstance(texts, list) or not all(isinstance(text, str) for text in texts):
        raise ValueError("texts must be a list of strings")
    try:
        inputs = tokenizer(texts, return_tensors="pt", truncation=True, padding=True, max_length=512)
        outputs = model(**inputs)
        probs = torch.nn.functional.softmax(outputs.logits, dim=-1)
        predicted_labels = torch.argmax(probs, dim=1).tolist()
        confidences = [probs[i, predicted_label].item() for i, predicted_label in enumerate(predicted_labels)]
        # 返回标签索引和置信度
        return [(predicted_label, confidence) for predicted_label, confidence in zip(predicted_labels, confidences)]
    except Exception as e:
        print(f"分析文本批次时出错: {e}")
        import traceback
        traceback.print_exc()
        return []


def main():
    # 关闭 oneDNN 自定义操作
    os.environ['TF_ENABLE_ONEDNN_OPTS'] = '0'
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

    # 输入输出文件路径
    input_file = r'D:\桌面\爬取数据\训练集_去重复.xlsx'

    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"读取文件 {input_file} 时出错: {e}")
        return

    if "text" not in df.columns or "label" not in df.columns:
        print("Excel 文件中没有 'text' 列或 'label' 列。请检查文件格式。")
        return

    # 过滤空值和非字符串
    df["text"] = df["text"].fillna("").astype(str)
    # 用 0 填充缺失值并转换为整数类型
    df["label"] = df["label"].fillna(0).astype(int)

    # 对文本列进行情感分析
    batch_size = 4  # 最小化 batch_size
    results = []
    true_labels = []
    for i in range(0, len(df["text"]), batch_size):
        batch_texts = df["text"][i:i + batch_size].tolist()
        # 过滤掉空字符串的批次
        batch_texts = [text for text in batch_texts if text.strip()]
        if not batch_texts:  # 如果批次为空，跳过
            print(f"跳过空批次 {i}-{i + batch_size}")
            continue
        try:
            batch_results = analyze_sentiment_batch(batch_texts, tokenizer, model)
            results.extend(batch_results)
            # 获取该批次的真实标签
            true_labels_batch = df["label"][i:i + batch_size].tolist()
            true_labels.extend(true_labels_batch)
        except Exception as e:
            print(f"处理批次 {i}-{i + batch_size} 时出错: {e}")
            import traceback
            traceback.print_exc()  # 打印详细的异常堆栈信息
            continue

    # 将分析结果添加到 DataFrame
    results_df = pd.DataFrame(results, columns=["LabelIndex", "Confidence"])
    df = pd.concat([df, results_df], axis=1)

    # 计算评估指标
    true_labels = df["label"].tolist()
    predicted_labels = df["LabelIndex"].tolist()
    accuracy = accuracy_score(true_labels, predicted_labels)
    precision = precision_score(true_labels, predicted_labels, average='weighted')
    recall = recall_score(true_labels, predicted_labels, average='weighted')
    f1 = f1_score(true_labels, predicted_labels, average='weighted')

    print(f"准确率: {accuracy}")
    print(f"精确率: {precision}")
    print(f"召回率: {recall}")
    print(f"F1 分数: {f1}")


if __name__ == "__main__":
    main()