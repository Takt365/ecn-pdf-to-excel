#!/usr/bin/env node
import fs from 'fs';
import path from 'path';
import { pathToFileURL } from 'url';
import XLSX from 'xlsx';
import { PDFParse } from 'pdf-parse';

const COLUMNS = [
  '直上品目コード',
  '直上品目テキスト',
  '親品目コード',
  '親品目テキスト'
];

function clean(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function isMaterialCode(value) {
  return /^(?:[A-Z]{1,3}\d{6,}[A-Z]?|\d{8,})$/.test(value);
}

function isItemCode(value) {
  return /^(?=.*\d)[A-Z0-9][A-Z0-9-]{5,}$/.test(value);
}

function parseRow(line) {
  const fields = clean(line).split(' ');
  if (fields.length < 7) return null;

  const materialIndex = fields.findIndex((field, index) => index >= 2 && isMaterialCode(field));
  if (materialIndex < 3) return null;

  const quantityIndex = fields.findIndex(
    (field, index) => index > materialIndex && /^\d+\.\d+$/.test(field)
  );
  if (quantityIndex < 0) return null;
  if (materialIndex >= quantityIndex) return null;

  const parentCodeIndex = isItemCode(fields[1]) ? 1 : 0;
  const parentCode = fields[parentCodeIndex];
  const parentText = fields.slice(parentCodeIndex + 1, materialIndex).join(' ');
  const upperCode = fields[materialIndex];
  const upperText = fields.slice(materialIndex + 1, quantityIndex).join(' ');
  if (!parentCode || !parentText || !upperCode || !upperText) return null;

  return {
    '直上品目コード': upperCode,
    '直上品目テキスト': upperText,
    '親品目コード': parentCode,
    '親品目テキスト': parentText
  };
}

export function parseReverseBom(text) {
  const rows = [];
  let headerFound = false;
  const lines = text
    .replace(/<PARSED TEXT FOR PAGE: \d+ \/ \d+>/g, '')
    .split(/\r?\n/);

  for (const rawLine of lines) {
    const line = clean(rawLine);
    if (!headerFound) {
      headerFound = line.includes('親機種コード') && line.includes('直上品目コード');
      continue;
    }
    const row = parseRow(line);
    if (row) rows.push(row);
  }
  return rows;
}

export async function convertReverseBom(pdfPath, outDir) {
  const base = path.basename(pdfPath, path.extname(pdfPath));
  fs.mkdirSync(outDir, { recursive: true });
  const parser = new PDFParse({ data: fs.readFileSync(pdfPath) });
  const parsed = await parser.getText();
  const rows = parseReverseBom(parsed.text);
  const outputPath = path.join(outDir, `${base}_机种清单.xlsx`);
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows, { header: COLUMNS });
  XLSX.utils.book_append_sheet(workbook, sheet, '机种清单');
  XLSX.writeFile(workbook, outputPath);
  return { outputPath, count: rows.length };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const pdfPath = process.argv[2]?.replace(/^['"]|['"]$/g, '');
  if (!pdfPath || !fs.existsSync(pdfPath)) {
    console.error('用法: node reverse_bom_to_excel.mjs <PDF路径> [--out <输出目录>]');
    process.exit(1);
  }

  const outIndex = process.argv.indexOf('--out');
  const outDir = outIndex >= 0 && process.argv[outIndex + 1]
    ? path.resolve(process.argv[outIndex + 1])
    : path.join(path.dirname(path.resolve(pdfPath)), 'output');
  const result = await convertReverseBom(pdfPath, outDir);
  console.log(`Excel: ${result.outputPath}`);
  console.log(`机种清单 ${result.count} 条`);
}
