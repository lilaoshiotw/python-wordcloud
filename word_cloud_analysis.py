import jieba
import pandas as pd
from wordcloud import WordCloud
import numpy as np
from PIL import Image
import matplotlib.pyplot as plt
from collections import Counter
import os
from docx import Document

def read_docx_file(file_path):
    """从Word文档读取内容"""
    try:
        doc = Document(file_path)
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        return '\n'.join(full_text)
    except Exception as e:
        print(f"读取文件时出错：{str(e)}")
        return ""

def process_text(text):
    """处理文本并统计词频"""
    # 定义停用词
    stop_words = set([
        '的', '了', '和', '是', '就', '都', '而', '及', '与', '着',
        '之', '在', '对', '等', '来', '用', '但', '并', '很', '又',
        '因为', '所以', '如果', '这个', '那个', '这样', '那样', '地',
        '得', '着', '说', '上', '下', '把', '让', '给', '到', '才',
        '已经', '可以', '会', '要', '去', '这', '那', '有', '个',
        '什么', '还', '能', '也', '没有', '就是', '我们', '你们',
        '他们', '她们', '它们', '自己', '这些', '那些', '一个', '一些',
        '\n', '\t', ' ', '，', '。', '！', '？', '；', '：', '"', '"',
        '（', '）', '【', '】', '[', ']', '、', '《', '》', '中',
        '为', '于', '其', '则', '由', '种', '式', '化', '时', '以',
        '使', '或', '被', '所', '一种', '一个', '通过', '进行', '该'
    ])

    # 使用jieba进行分词
    words = jieba.cut(text)
    
    # 过滤停用词和单个字符
    valid_words = [word for word in words if word not in stop_words and len(word) > 1]
    
    # 统计词频
    word_freq = Counter(valid_words)
    
    return word_freq

def generate_word_cloud(word_freq, output_path):
    """生成词云图片"""
    # 读取蒙版图片
    mask = np.array(Image.open('mask.jpeg'))
    
    # 使用系统字体
    font_path = 'C:\\Windows\\Fonts\\simhei.ttf'  # Windows系统黑体字体路径
    
    # 创建WordCloud对象
    wc = WordCloud(
        font_path=font_path,  # 使用系统黑体字体
        background_color='white',
        mask=mask,
        max_words=200,
        width=1000,
        height=800,
        contour_width=1,  # 添加轮廓
        contour_color='black'  # 轮廓颜色
    )
    
    # 生成词云
    wc.generate_from_frequencies(word_freq)
    
    # 保存词云图片
    plt.figure(figsize=(10, 8))
    plt.imshow(wc, interpolation='bilinear')
    plt.axis('off')
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

def save_word_freq(word_freq, output_path):
    """保存词频统计到Excel"""
    df = pd.DataFrame(word_freq.items(), columns=['词语', '频次'])
    df = df.sort_values('频次', ascending=False)
    df.to_excel(output_path, index=False)

def main():
    # 指定Word文档路径
    input_file = '梁媛古筝协奏曲《海之波澜》音乐与演奏分析.docx'
    
    if not os.path.exists(input_file):
        print(f"找不到文件：{input_file}")
        return
    
    # 读取文档内容
    print("正在读取Word文档...")
    content = read_docx_file(input_file)
    
    if not content:
        print("文档内容为空，请检查文件")
        return
    
    # 处理文本并统计词频
    print("正在处理文本...")
    word_freq = process_text(content)
    
    # 生成词云
    print("正在生成词云...")
    generate_word_cloud(word_freq, 'wordcloud.png')
    
    # 保存词频统计
    print("正在保存词频统计...")
    save_word_freq(word_freq, 'word_frequency.xlsx')
    
    print("处理完成！")
    print(f"词云图片已保存为：wordcloud.png")
    print(f"词频统计已保存为：word_frequency.xlsx")

if __name__ == '__main__':
    main() 