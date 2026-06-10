import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import type { ArtifactFormat, ContentBlock, CreateArtifactRequest, NormalizedDocument } from '../types/artifacts';
import { logger } from './logger';

export interface RenderedFile {
  format: ArtifactFormat;
  filename: string;
  contentType: string;
  buffer: Buffer;
}

function filenameSafe(value: string): string {
  const cleaned = value
    .trim()
    .replace(/[^a-zA-Z0-9._ -]+/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .slice(0, 80);
  return cleaned || 'document';
}

function stripMarkdownInline(value: string): string {
  return value
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/__(.*?)__/g, '$1')
    .replace(/\*(.*?)\*/g, '$1')
    .replace(/_(.*?)_/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$1 ($2)');
}

function isDividerRow(cells: string[]): boolean {
  return cells.length > 0 && cells.every((cell) => /^:?-{3,}:?$/.test(cell.trim()));
}

function parseTable(lines: string[], startIndex: number): { block: ContentBlock; nextIndex: number } | undefined {
  const rows: string[][] = [];
  let index = startIndex;
  while (index < lines.length && lines[index].includes('|')) {
    const cells = lines[index]
      .split('|')
      .map((cell) => stripMarkdownInline(cell.trim()))
      .filter((cell, cellIndex, allCells) => !(cell === '' && (cellIndex === 0 || cellIndex === allCells.length - 1)));
    if (!isDividerRow(cells)) rows.push(cells);
    index += 1;
  }
  if (rows.length < 1) return undefined;
  return { block: { type: 'table', rows }, nextIndex: index };
}

export function normalizeDocument(input: CreateArtifactRequest): NormalizedDocument {
  const style = input.style || 'plain';
  if (input.blocks?.length) {
    return {
      title: input.title,
      style,
      blocks: input.blocks,
      sourceNotes: input.sourceNotes,
    };
  }

  const markdown = input.markdown || '';
  const lines = markdown.replace(/\r\n/g, '\n').split('\n');
  const blocks: ContentBlock[] = [];
  let paragraph: string[] = [];
  let index = 0;

  const flushParagraph = () => {
    if (paragraph.length === 0) return;
    blocks.push({ type: 'paragraph', text: stripMarkdownInline(paragraph.join(' ')) });
    paragraph = [];
  };

  while (index < lines.length) {
    const line = lines[index].trim();
    if (!line) {
      flushParagraph();
      index += 1;
      continue;
    }

    const heading = /^(#{1,3})\s+(.+)$/.exec(line);
    if (heading) {
      flushParagraph();
      blocks.push({
        type: 'heading',
        level: heading[1].length as 1 | 2 | 3,
        text: stripMarkdownInline(heading[2]),
      });
      index += 1;
      continue;
    }

    if (line === '---' || line === '***') {
      flushParagraph();
      blocks.push({ type: 'pageBreak' });
      index += 1;
      continue;
    }

    if (line.includes('|')) {
      const table = parseTable(lines, index);
      if (table) {
        flushParagraph();
        blocks.push(table.block);
        index = table.nextIndex;
        continue;
      }
    }

    const listItems: string[] = [];
    const ordered = /^\d+\.\s+/.test(line);
    while (index < lines.length) {
      const current = lines[index].trim();
      const unorderedMatch = /^[-*]\s+(.+)$/.exec(current);
      const orderedMatch = /^\d+\.\s+(.+)$/.exec(current);
      if (ordered && orderedMatch) {
        listItems.push(stripMarkdownInline(orderedMatch[1]));
      } else if (!ordered && unorderedMatch) {
        listItems.push(stripMarkdownInline(unorderedMatch[1]));
      } else {
        break;
      }
      index += 1;
    }
    if (listItems.length > 0) {
      flushParagraph();
      blocks.push({ type: 'list', ordered, items: listItems });
      continue;
    }

    paragraph.push(line);
    index += 1;
  }

  flushParagraph();
  return {
    title: input.title,
    style,
    blocks: blocks.length > 0 ? blocks : [{ type: 'paragraph', text: '' }],
    sourceNotes: input.sourceNotes,
  };
}

function headingLevel(level: 1 | 2 | 3 | undefined) {
  if (level === 1) return HeadingLevel.HEADING_1;
  if (level === 2) return HeadingLevel.HEADING_2;
  return HeadingLevel.HEADING_3;
}

function docxChildren(document: NormalizedDocument): Array<Paragraph | Table> {
  const children: Array<Paragraph | Table> = [
    new Paragraph({
      text: document.title,
      heading: document.style === 'plain' ? HeadingLevel.TITLE : HeadingLevel.HEADING_1,
      alignment: document.style === 'letter' ? AlignmentType.LEFT : AlignmentType.CENTER,
      spacing: { after: 300 },
    }),
  ];

  for (const block of document.blocks) {
    if (block.type === 'heading') {
      children.push(new Paragraph({ text: block.text || '', heading: headingLevel(block.level), spacing: { before: 240, after: 120 } }));
    } else if (block.type === 'paragraph') {
      children.push(new Paragraph({ children: [new TextRun(block.text || '')], spacing: { after: 180 } }));
    } else if (block.type === 'list') {
      for (const item of block.items || []) {
        children.push(new Paragraph({ text: item, bullet: block.ordered ? undefined : { level: 0 }, numbering: block.ordered ? { reference: 'default-numbering', level: 0 } : undefined, spacing: { after: 90 } }));
      }
    } else if (block.type === 'table') {
      children.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: (block.rows || []).map((row) => new TableRow({
          children: row.map((cell) => new TableCell({
            children: [new Paragraph({ children: [new TextRun(cell)] })],
          })),
        })),
      }));
    } else if (block.type === 'pageBreak') {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }
  }
  return children;
}

async function renderDocx(document: NormalizedDocument): Promise<RenderedFile> {
  const doc = new Document({
    numbering: {
      config: [
        {
          reference: 'default-numbering',
          levels: [
            {
              level: 0,
              format: 'decimal',
              text: '%1.',
              alignment: AlignmentType.LEFT,
            },
          ],
        },
      ],
    },
    sections: [{ children: docxChildren(document) }],
  });
  const buffer = await Packer.toBuffer(doc);
  return {
    format: 'docx',
    filename: `${filenameSafe(document.title)}.docx`,
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer,
  };
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

export function renderDocumentHtml(document: NormalizedDocument): string {
  const body = document.blocks.map((block) => {
    if (block.type === 'heading') return `<h${block.level || 3}>${escapeHtml(block.text || '')}</h${block.level || 3}>`;
    if (block.type === 'paragraph') return `<p>${escapeHtml(block.text || '')}</p>`;
    if (block.type === 'list') {
      const tag = block.ordered ? 'ol' : 'ul';
      return `<${tag}>${(block.items || []).map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</${tag}>`;
    }
    if (block.type === 'table') {
      return `<table>${(block.rows || []).map((row, rowIndex) => `<tr>${row.map((cell) => rowIndex === 0 ? `<th>${escapeHtml(cell)}</th>` : `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`).join('')}</table>`;
    }
    return '<div class="page-break"></div>';
  }).join('\n');

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(document.title)}</title>
  <style>
    @page { margin: 0.75in; }
    body { color: #111827; font-family: Georgia, 'Times New Roman', serif; font-size: 12pt; line-height: 1.45; }
    h1 { font-size: 22pt; text-align: center; margin: 0 0 24pt; }
    h2 { font-size: 16pt; margin-top: 18pt; }
    h3 { font-size: 13pt; margin-top: 14pt; }
    p { margin: 0 0 10pt; }
    table { border-collapse: collapse; margin: 12pt 0; width: 100%; }
    th, td { border: 1px solid #d1d5db; padding: 6pt; text-align: left; vertical-align: top; }
    th { background: #f3f4f6; font-weight: 700; }
    .page-break { page-break-after: always; }
  </style>
</head>
<body>
  <h1>${escapeHtml(document.title)}</h1>
  ${body}
</body>
</html>`;
}

function simplePdfBuffer(document: NormalizedDocument): Buffer {
  const text = [document.title, ...document.blocks.flatMap((block) => {
    if (block.type === 'heading' || block.type === 'paragraph') return [block.text || ''];
    if (block.type === 'list') return block.items || [];
    if (block.type === 'table') return (block.rows || []).map((row) => row.join(' | '));
    return [''];
  })].join('\n');
  const escaped = text.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
  const content = `BT /F1 12 Tf 72 720 Td 14 TL (${escaped.slice(0, 3500).replace(/\n/g, ') Tj T* (')}) Tj ET`;
  const objects = [
    '1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj',
    '2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj',
    '3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >> endobj',
    `4 0 obj << /Length ${Buffer.byteLength(content)} >> stream\n${content}\nendstream endobj`,
    '5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj',
  ];
  let pdf = '%PDF-1.4\n';
  const offsets: number[] = [0];
  for (const object of objects) {
    offsets.push(Buffer.byteLength(pdf));
    pdf += `${object}\n`;
  }
  const xrefOffset = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets.slice(1)) pdf += `${String(offset).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer << /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return Buffer.from(pdf);
}

async function renderPdf(document: NormalizedDocument): Promise<RenderedFile> {
  try {
    const { chromium } = await import('playwright');
    const browser = await chromium.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    try {
      const page = await browser.newPage();
      await page.setContent(renderDocumentHtml(document), { waitUntil: 'load' });
      const buffer = await page.pdf({ format: 'Letter', printBackground: true });
      return {
        format: 'pdf',
        filename: `${filenameSafe(document.title)}.pdf`,
        contentType: 'application/pdf',
        buffer,
      };
    } finally {
      await browser.close();
    }
  } catch (error) {
    logger.warn('Playwright PDF rendering unavailable, using simple PDF fallback', {
      error: error instanceof Error ? error.message : String(error),
    });
    return {
      format: 'pdf',
      filename: `${filenameSafe(document.title)}.pdf`,
      contentType: 'application/pdf',
      buffer: simplePdfBuffer(document),
    };
  }
}

export async function renderDocument(input: CreateArtifactRequest): Promise<{ document: NormalizedDocument; files: RenderedFile[] }> {
  const document = normalizeDocument(input);
  const formats = Array.from(new Set(input.formats));
  const files: RenderedFile[] = [];
  for (const format of formats) {
    if (format === 'docx') files.push(await renderDocx(document));
    if (format === 'pdf') files.push(await renderPdf(document));
  }
  return { document, files };
}
