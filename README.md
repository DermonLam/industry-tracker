# Industry Tracker - 行业信息跟踪抓取

分析文章（特别是微信公众号文章）内容，按用户指定的维度提取关键信息，生成结构化的 Word 格式分析报告。

## 功能

- 逐篇获取文章全文内容
- 多篇文章交叉去重
- 按维度分类提取（政策、投资并购、战略、监管、重点公司动态等）
- 生成格式化 Word 文档（.docx）
- 默认物业行业配置，支持自定义行业和维度

## 安装

### Node.js 依赖

```bash
cd scripts
npm install
```

### Python 依赖

```bash
pip install python-docx
```

两种运行时选一种即可，Node.js 优先。

## 使用

### 命令行

```bash
# Node.js
node scripts/generate_report.js --input data.json --output report.docx

# Python
python scripts/generate_report.py --input data.json --output report.docx

# 也可以从 stdin 传入 JSON
echo '<JSON>' | node scripts/generate_report.js
```

### JSON 输入格式

```json
{
  "outputFile": "article_analysis_report.docx",
  "title": "【物业行业动态分析】2025-06-14",
  "source": "来源：XX公众号",
  "sourceUrl": "https://...",
  "date": "2025-06-14",
  "sections": [
    {
      "heading": "一、政策动态",
      "type": "bullets",
      "items": ["条目一", "条目二"]
    },
    {
      "heading": "五、重点公司动态",
      "type": "paragraphs",
      "subsections": [
        { "subheading": "1. 碧桂园服务", "paragraphs": ["动态内容"] }
      ]
    }
  ]
}
```

支持的 section 类型：
- `bullets`：项目符号列表
- `table`：表格（需提供 `columns` 和 `rows`）
- `paragraphs`：带子标题的段落

### 在 WorkBuddy 中安装

将此仓库克隆到 WorkBuddy 的 skills 目录：

```bash
git clone https://github.com/DermonLam/industry-tracker.git \
  ~/.workbuddy/skills/industry-tracker
```

安装后，当你在对话中提到"物业周报整理"、"物业文章整理"等关键词并提供文章链接时，WorkBuddy 会自动激活此 skill。

## 文档格式

- 字体：Arial
- Heading 1：深蓝 #1F3864
- Heading 2：蓝 #2E75B6
- Heading 3：深绿 #375623
- 页面：A4，边距约 2.1cm
- 页眉：报告标题
- 页脚：页码居中

## License

MIT
