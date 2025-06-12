# 文档词云分析工具

这是一个用于分析文档内容并生成词云的Python工具。支持Word文档（.docx）格式的输入，可以生成美观的中国地图形状词云。

## 功能特点

- 支持读取Word文档（.docx格式）
- 使用结巴分词进行中文分词
- 生成中国地图形状的词云图片
- 输出词频统计Excel文件
- 支持自定义停用词过滤

## 环境要求

- Python 3.6+
- 依赖包：
  - jieba：中文分词
  - wordcloud：词云生成
  - python-docx：Word文档处理
  - numpy：数据处理
  - pandas：数据分析
  - Pillow：图像处理
  - matplotlib：图像绘制
  - openpyxl：Excel文件处理

## 安装步骤

1. 克隆或下载本项目
2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

1. 准备Word文档：
   - 将需要分析的文档放在项目根目录下
   - 确保文档为.docx格式

2. 运行脚本：
```bash
python word_cloud_analysis.py
```

## 输出文件

- `wordcloud.png`：生成的词云图片（中国地图形状）
- `word_frequency.xlsx`：词频统计Excel文件（按频次降序排列）

## 自定义设置

- 词云形状：使用 `mask.jpeg` 文件定义
- 停用词：可在代码中的 `stop_words` 集合中添加或删除
- 字体：默认使用系统的黑体字体（simhei.ttf）

## 注意事项

- 确保Word文档可以正常打开和读取
- 可以根据需要调整停用词列表
- 词云的清晰度和大小可以通过代码中的参数调整 