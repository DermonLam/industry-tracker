# Article Report Generator

分析文章（特别是微信公众号文章）内容，按用户指定的维度提取关键信息，生成结构化的Word格式分析报告。

## 触发规则

当用户说出以下任一类型语句并附带一个或多个文章链接时，直接激活此 skill：

- "物业周报整理" / "物业文章整理" / "物业行业信息整理"
- "整理物业文章" / "物业动态汇总" / "物业周报"
- "分析这些物业文章" / "物业信息整理"
- 其他表述满足核心触发模式：**物业 + (整理/汇总/周报/分析) + URL 列表**

若用户未提及"物业"但提供了文章链接并要求"按维度整理成Word"，同样触发此 skill，但使用空白默认配置——需向用户确认维度和公司清单后再执行。

---

## 默认配置

以下为物业行业默认配置，用户可在指令中覆盖：

| 配置项 | 默认值 |
|--------|--------|
| 行业 | 物业 |
| 维度 | 政策、投资并购、战略、监管、重点公司动态 |
| 公司清单 | 碧桂园服务、万物云、绿城服务、招商积余、华润万象生活、保利物业、中海物业、雅生活服务、世茂服务 |
| 交付渠道 | 微信 clawbot + 163 邮箱（两者都发） |
| 文档标题格式 | 【物业行业动态分析】YYYY-MM-DD |

**配置来源优先级**：用户当前指令 > 项目 MEMORY.md 中的行业配置 > 上表默认值

**用户覆盖示例**：
- "按XX维度整理" → 替换维度列表
- "重点关注XX公司" → 替换/增补公司清单
- "只发邮件" / "先别发微信" → 调整交付渠道
- "XX行业文章整理" → 替换行业及关联默认值

---

## 工作流程

### Step 1: 逐篇获取文章内容

对用户提供的每个 URL 使用 WebFetch 逐一获取全文内容。

**无论1篇还是N篇，统一逐篇获取，每篇读取后立即标注来源 URL。**

WebFetch prompt 指南：
```
提取这篇文章的完整正文内容，包括标题、作者、发布时间等元数据。请尽可能保留所有文本内容。
```

### Step 2: 跨文章去重

**单篇文章**：跳过此步骤。

**多篇文章**：对所有文章内容进行交叉比对，识别重复报道事项。

**去重判定标准**：同一主体 + 同一事件 + 同一时间 = 重复项。示例：
- "碧桂园服务6月5日回购80万股" 在两篇文章中均出现 → 只保留一次
- "住建部发布公积金条例修订征求意见稿" 在多篇文章重复 → 只保留一次，标注多篇文章来源

**去重原则**：
- 完全相同的事项只输出一次，在来源中列出所有提及该事项的文章
- 同一事项但细节不同（如A文说回购80万股，B文说回购80万股+额外数据）→ 合并为一条，取最完整的描述
- 同一公司不同事项 → 不合并，分别列出

去重时在内部完成交叉比对并记录结果，确认后再进入分类步骤。

### Step 3: 分类提取

根据配置中的维度列表，对（去重后的）文章内容进行分类提取。

| 维度 | 说明 |
|------|------|
| 政策（Policy） | 政府政策、法规、规划、统计制度等 |
| 投资并购（Investment/M&A） | 股权交易、投资、股份回购、增持等 |
| 战略（Strategy） | 企业战略合作、业务拓展、中标项目、新业务布局等 |
| 监管（Regulation） | 行业协会倡议、团体标准、监管要求、规范文件等 |
| 重点公司动态 | 用户指定的特定公司相关动态 |

分类原则：
- 只提取与指定维度相关的内容
- 保持客观，不添加主观解读
- 每条信息包含关键要素：主体、动作、时间、核心数据
- 若文章未提及某维度内容，明确标注"本文未提及相关动态"

**重点公司清单处理**：对配置中的每家公司逐一检查，无动态则标注"本文未提及XX相关动态"。

### Step 4: 生成 Word 文档

使用 docx-js（Node.js）或 python-docx（Python）生成 Word 文档。两套脚本功能等价，接受相同的 JSON 输入。

#### 运行时选择

1. 优先检测 Node.js 是否可用 → 使用 `scripts/generate_report.js`
2. 退而使用 Python → 使用 `scripts/generate_report.py`
3. 两者都不可用 → 报错提示安装

#### 调用方式

```bash
# Node.js
echo '<JSON>' | node scripts/generate_report.js
# 或
node scripts/generate_report.js --input data.json --output report.docx

# Python
python scripts/generate_report.py --input data.json --output report.docx
```

#### JSON 输入格式

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
        { "subheading": "1. 碧桂园服务", "paragraphs": ["本文未提及碧桂园服务相关动态。"] }
      ]
    }
  ]
}
```

#### 文档格式规范

- 字体：Arial（全文档统一）
- 标题颜色：Heading 1 深蓝 `#1F3864`，Heading 2 蓝 `#2E75B6`，Heading 3 深绿 `#375623`
- 页眉：分析报告标题
- 页脚：页码居中
- 页面：A4，边距 1200 DXA（约2.1cm）

### Step 5: 交付

交付选项（用户未指定时默认全选 A+B）：

| 选项 | 方式 | 说明 |
|------|------|------|
| A | `deliver_attachments` → 微信 clawbot | 将 .docx 文件发送到用户手机微信 |
| B | netease-mail → 163 邮箱 | 发送邮件，主题格式：`【物业行业动态分析】YYYY-MM-DD`，正文简要列出各维度关键条目，.docx 作为附件 |
| C | 仅生成文件 | 不发送任何渠道，文件留在本地 |

**用户控制示例**：
- "只发邮件" → 仅执行 B
- "先别发微信" → 仅执行 B
- "只生成文件" → 仅执行 C
- 未说明 → 默认 A + B

#### 5A. 发送到微信 clawbot

调用 `deliver_attachments` 工具，将生成的 `.docx` 文件发送到用户手机微信（clawbot）。

> 若用户未开启 WorkBuddy 小程序的"产物回传到小程序"设置，文件可能无法送达。此时提示用户前往小程序开启。

#### 5B. 发送到 163 邮箱

1. 加载 netease-mail skill（`Skill` 工具，skill 参数为 `netease-mail`）
2. 按 skill 指引发送邮件
3. 若 netease-mail 连接器未连接，提示用户连接后再重试

#### 5C. 交付完成确认

文件发送后，向用户确认交付结果（如"已通过 clawbot 发送到微信 + 163 邮箱各一份"），并简要列出报告内容摘要（各维度各几条）。

---

## 关键技术说明

### docx-js 使用要点

- 安装：`npm install docx`（在项目目录或 workspace）
- 运行：`node generate_report.js`
- 列表必须使用 `LevelFormat.BULLET` + numbering config，禁止手动插入 bullet 字符
- 表格需同时设置 `columnWidths`（表格级）和 `width`（单元格级）
- 分页用 `new Paragraph({ children: [new PageBreak()] })`

### python-docx 使用要点

- 安装：`pip install python-docx`
- 运行：`python generate_report.py --input data.json --output report.docx`
- 中文字符无需特殊处理

### 中文字符处理

docx-js 默认支持 Unicode，中文字符无需特殊处理。确保脚本文件保存为 UTF-8 编码。

## 参考资料

- docx-js 官方文档：https://docx.js.org/
- python-docx 官方文档：https://python-docx.readthedocs.io/
