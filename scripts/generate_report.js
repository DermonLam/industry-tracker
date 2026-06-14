/**
 * 文章分析报告 Word 文档生成脚本（docx-js）
 *
 * 用法：
 *   echo '<JSON>' | node generate_report.js
 *   node generate_report.js --input data.json
 *   node generate_report.js --input data.json --output report.docx
 *
 * JSON 输入格式见 SKILL.md
 *
 * 依赖：npm install docx
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun,
  HeadingLevel, AlignmentType, LevelFormat,
  BorderStyle, WidthType, ShadingType,
  Table, TableRow, TableCell,
  PageBreak, PageNumber, Header, Footer
} = require('docx');

// ============================================================
// 读取输入
// ============================================================

function loadInput() {
  const args = process.argv.slice(2);
  let inputFile = null;
  let outputFile = null;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--input' && args[i + 1]) inputFile = args[++i];
    else if (args[i] === '--output' && args[i + 1]) outputFile = args[++i];
  }

  let jsonStr;

  if (inputFile) {
    jsonStr = fs.readFileSync(inputFile, 'utf-8');
  } else {
    // 从 stdin 读取
    const chunks = [];
    const stdin = process.stdin;
    if (stdin.isTTY) {
      console.error('用法: echo \'<JSON>\' | node generate_report.js');
      console.error('      node generate_report.js --input data.json [--output report.docx]');
      process.exit(1);
    }
    jsonStr = fs.readFileSync('/dev/stdin', 'utf-8');
  }

  const data = JSON.parse(jsonStr);
  data.outputFile = outputFile || data.outputFile || 'article_analysis_report.docx';
  return data;
}

// ============================================================
// 文档构建
// ============================================================

const A4_WIDTH = 11906;
const A4_HEIGHT = 16838;
const MARGIN = 1200;
const CONTENT_WIDTH = A4_WIDTH - MARGIN * 2;  // ~9506 DXA

function buildChildren(data) {
  const kids = [];

  // 标题
  kids.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    children: [new TextRun(data.title || '文章分析报告')]
  }));
  kids.push(spacer());

  // 来源信息
  if (data.source) {
    kids.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: data.source, size: 20, color: '666666', font: 'Arial' })]
    }));
  }
  if (data.sourceUrl) {
    kids.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `原文链接：${data.sourceUrl}`, size: 18, color: '888888', font: 'Arial' })]
    }));
  }
  // 多来源链接
  if (data.sourceUrls && data.sourceUrls.length > 0) {
    for (const s of data.sourceUrls) {
      kids.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `${s.title || ''} - ${s.url}`, size: 18, color: '888888', font: 'Arial' })]
      }));
    }
  }
  kids.push(spacer());

  // 各章节
  for (const sec of (data.sections || [])) {
    // Heading 2
    kids.push(new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun(sec.heading)]
    }));

    if (sec.type === 'bullets') {
      for (const item of (sec.items || [])) {
        kids.push(bullet(item));
      }
    } else if (sec.type === 'table') {
      kids.push(buildTable(sec));
    } else if (sec.type === 'paragraphs') {
      for (const sub of (sec.subsections || [])) {
        // Heading 3
        kids.push(new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun(sub.subheading)]
        }));
        for (const para of (sub.paragraphs || [])) {
          kids.push(new Paragraph({
            children: [new TextRun({ text: para, font: 'Arial' })]
          }));
        }
        kids.push(spacer());
      }
    }

    kids.push(spacer());
  }

  // 结尾
  kids.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: '—— 报告结束 ——', size: 20, color: 'AAAAAA', font: 'Arial' })]
  }));

  return kids;
}

function buildDocument(children, title) {
  return new Document({
    styles: {
      default: {
        document: { run: { font: 'Arial', size: 22 } }
      },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: '1F3864' },
          paragraph: { spacing: { before: 300, after: 150 } }
        },
        {
          id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: '2E75B6' },
          paragraph: { spacing: { before: 200, after: 100 } }
        },
        {
          id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 22, bold: true, font: 'Arial', color: '375623' },
          paragraph: { spacing: { before: 150, after: 80 } }
        }
      ]
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: '\u2022',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }]
        }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: A4_WIDTH, height: A4_HEIGHT },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: title || '文章分析报告', size: 18, color: '888888', font: 'Arial' })]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: '第 ', size: 18, color: '888888', font: 'Arial' }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '888888', font: 'Arial' }),
              new TextRun({ text: ' 页', size: 18, color: '888888', font: 'Arial' })
            ]
          })]
        })
      },
      children
    }]
  });
}

// ---- 辅助函数 ----

function spacer() {
  return new Paragraph({ children: [new TextRun('')] });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: [new TextRun({ text, font: 'Arial' })]
  });
}

function buildTable(sec) {
  const colWidth = Math.floor(CONTENT_WIDTH / sec.columns.length);
  const rows = [
    new TableRow({
      children: sec.columns.map(c => cell(typeof c === 'string' ? c : c.header, true, colWidth))
    }),
    ...sec.rows.map(row =>
      new TableRow({
        children: row.map((c, i) => cell(typeof c === 'string' ? c : c.text, false, colWidth))
      })
    )
  ];
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: sec.columns.map(() => colWidth),
    rows
  });
}

function cell(text, isHeader, width) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  return new TableCell({
    borders: { top: border, bottom: border, left: border, right: border },
    width: { size: width, type: WidthType.DXA },
    shading: isHeader ? { fill: 'D5E8F0', type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, bold: isHeader, font: 'Arial' })]
    })]
  });
}

// ============================================================
// 主入口
// ============================================================

function main() {
  const data = loadInput();
  const children = buildChildren(data);
  const doc = buildDocument(children, data.title);

  Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(data.outputFile, buffer);
    console.log(`✅ 文档已生成：${data.outputFile}`);
  }).catch(err => {
    console.error('❌ 生成失败：', err);
    process.exit(1);
  });
}

main();
