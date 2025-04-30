import pandas as pd
import jieba
import jieba.posseg as pseg
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.decomposition import LatentDirichletAllocation
from gensim.models import CoherenceModel
from gensim.corpora import Dictionary
import numpy as np
import joblib
import os
from collections import Counter
from wordcloud import WordCloud

# 加载哈工大停用词库（增强版）
def load_stopwords(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            stopwords = [line.strip() for line in f.readlines()]
        # 新增通用中文停用词
        extra_stopwords = {'的', '了', '是', '就', '也', '都', '吗', '啊', '这', '在', '我', '你', '他', '她', '它', '不', '没有', '说',
                           '要', '就', '那', '有'}
        return list(set(stopwords) | extra_stopwords)
    except FileNotFoundError:
        print(f"错误：未找到停用词文件 {file_path}")
        return []

# 自定义词典
custom_words = {
    '饥饿营销': 'n',
    '快闪': 'n',
    '通贩': 'v',
    '补货': 'v',
    '五百款': 'n',
    '对角巷': 'n',
    '海德薇': 'n',
    '哈利波特': 'n',
    '不太好看': 'a',
    '赫奇帕奇': 'n',
    '名创优品': 'n'
}

# 数据预处理函数
def preprocess_text(text, stopwords):
    words_with_pos = pseg.cut(text)
    meaningful_pos = ['n', 'v', 'a']
    filtered_words = []
    for word, pos in words_with_pos:
        if pos[0] in meaningful_pos and word not in stopwords:
            filtered_words.append(word)
    return filtered_words

# 新增主题代表性文本展示
def show_representative_texts(df, topic_probs, text_column, top_n=5):
    df['topic_prob'] = np.max(topic_probs, axis=1)
    for topic in range(topic_probs.shape[1]):
        print(f"\n主题 {topic} 代表性评论：")
        samples = df[df['topic'] == topic].nlargest(top_n, 'topic_prob')
        for idx, row in samples.iterrows():
            print(f"[概率 {row.topic_prob:.2f}] {str(row[text_column]).strip()[:60]}...")

# 输出主题关键词
def print_topic_keywords(lda_model, vectorizer, n_top_words):
    feature_names = vectorizer.get_feature_names_out()
    for topic_idx, topic in enumerate(lda_model.components_):
        top_words_idx = topic.argsort()[:-n_top_words - 1:-1]
        top_words = [feature_names[i] for i in top_words_idx]
        print(f"主题 {topic_idx}: {' '.join(top_words)}")

# 生成词云
def generate_wordcloud(word_counter, save_path):
    wordcloud = WordCloud(font_path='simhei.ttf', background_color='white').generate_from_frequencies(word_counter)
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    wordcloud.to_file(os.path.join(save_path, 'wordcloud.png'))
    print(f"词云图已保存到 {os.path.join(save_path, 'wordcloud.png')}")

# 训练和评估 LDA 模型
def train_and_evaluate_lda(X, processed_texts, dictionary, vectorizer, min_topics, max_topics):
    best_lda_model = None
    best_n_topics = 0
    best_coherence = -1
    perplexity_values = []
    coherence_values = []
    topn = 20  # 每个主题取前 20 个词来计算一致性

    for n_topics in range(min_topics, max_topics + 1):
        lda_model = LatentDirichletAllocation(n_components=n_topics, random_state=42)
        lda_model.fit(X)

        # 计算困惑度
        perplexity = lda_model.perplexity(X)
        perplexity_values.append(perplexity)

        # 计算主题一致性
        corpus = [dictionary.doc2bow(text) for text in processed_texts]
        topic_word_dist = lda_model.components_ / lda_model.components_.sum(axis=1)[:, np.newaxis]

        # 将主题 - 词分布矩阵转换为每个主题的前 topn 个词的 ID 列表
        topics = []
        for topic in topic_word_dist:
            top_words_idx = topic.argsort()[:-topn - 1:-1]
            topics.append(top_words_idx.tolist())

        coherence_model = CoherenceModel(topics=topics, texts=processed_texts, dictionary=dictionary, corpus=corpus, coherence='c_v')
        coherence = coherence_model.get_coherence()
        coherence_values.append(coherence)

        if coherence > best_coherence:
            best_lda_model = lda_model
            best_n_topics = n_topics
            best_coherence = coherence

        print(f"主题数量: {n_topics}, 困惑度: {perplexity:.4f}, 主题一致性: {coherence:.4f}")

    return best_lda_model, best_n_topics, best_coherence, perplexity_values, coherence_values

def main():
    # 初始化Jieba并加载词典
    jieba.initialize()
    for word, pos in custom_words.items():
        jieba.add_word(word, tag=pos)

    # 加载停用词
    stopwords = load_stopwords(r'D:\桌面\优衣库主题分析\哈工大停用词表.txt')
    if not stopwords:
        return

    # 读取数据
    file_path = r'D:\桌面\爬取数据\数据集结果\名创优品\名创优品消极.xlsx'
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"错误：未找到 Excel 文件 {file_path}")
        return

    text_column = 'text'
    print("正在预处理文本...")
    df['processed_text'] = df[text_column].apply(lambda x: preprocess_text(str(x), stopwords))

    # 词袋模型向量化
    vectorizer = CountVectorizer()
    X = vectorizer.fit_transform([' '.join(doc) for doc in df['processed_text']])

    # 准备计算主题一致性的数据
    dictionary = Dictionary(df['processed_text'])

    # 设置主题数量范围
    min_topics, max_topics = 2, 8

    # 训练模型
    print("正在训练模型...")
    best_lda_model, best_n_topics, best_coherence, perplexity_values, coherence_values = train_and_evaluate_lda(
        X, df['processed_text'], dictionary, vectorizer, min_topics, max_topics)

    print(f"\n最佳主题数量: {best_n_topics}, 最佳主题一致性: {best_coherence:.4f}")

    # 输出主题关键词
    print_topic_keywords(best_lda_model, vectorizer, 15)

    # 新增代表性文本展示
    topic_probs = best_lda_model.transform(X)
    df['topic'] = np.argmax(topic_probs, axis=1)
    show_representative_texts(df, topic_probs, text_column)

    # 保存模型
    save_path = r'D:\桌面\名创优品主题分析\名创优品积极消极主题'
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    model_file_path = os.path.join(save_path, 'best_lda_model名创优品DS消极.pkl')
    joblib.dump(best_lda_model, model_file_path)
    print(f"\n最佳模型已保存到 {model_file_path}")

    # 生成词云
    all_words = [word for sublist in df['processed_text'] for word in sublist]
    generate_wordcloud(Counter(all_words), save_path)

    # 保存结果
    new_file_path = r'D:\桌面\爬取数据\数据集结果\名创优品\lda模型结果\积极主题\名创优品_消极_DS主题.xlsx'
    df.to_excel(new_file_path, index=False)
    print(f"\n分析结果已保存到 {new_file_path}")

    return best_lda_model

if __name__ == "__main__":
    result = main()