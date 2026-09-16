#!/usr/bin/env node
import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';
import { PDFParse } from 'pdf-parse';
import { parseReverseBom } from './reverse_bom_to_excel.mjs';

const COLUMNS = [
  '直上品目コード',
  '直上品目テキスト',
  '親品目コード',
  '親品目テキスト'
];

function cleanArg(value) {
  return value?.replace(/^['"]|['"]$/g, '');
}

function ensureOutputDir(outputPath) {
  const outputDir = path.dirname(outputPath);
  if (fs.existsSync(outputDir)) {
    console.log(`输出目录已存在，跳过创建: ${outputDir}`);
    return;
  }
  fs.mkdirSync(outputDir, { recursive: true });
  console.log(`已创建输出目录: ${outputDir}`);
}

async function readPdfRows(filePath) {
  const parser = new PDFParse({ data: fs.readFileSync(filePath) });
  const parsed = await parser.getText();
  return parseReverseBom(parsed.text);
}

export async function mergeReverseBom(inputDir, outputPath) {
  const files = fs.readdirSync(inputDir)
    .filter(name => name.toLowerCase().endsWith('.pdf'))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  if (files.length === 0) {
    throw new Error(`目录中没有找到 PDF 文件: ${inputDir}`);
  }

  const rows = [];
  for (const file of files) {
    const fileRows = await readPdfRows(path.join(inputDir, file));
    rows.push(...fileRows);
  }

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows, { header: COLUMNS });
  XLSX.utils.book_append_sheet(workbook, sheet, '机种清单');
  ensureOutputDir(outputPath);
  XLSX.writeFile(workbook, outputPath);
  return { files: files.length, rows: rows.length, outputPath };
}

const inputDir = path.resolve(cleanArg(process.argv[2]) || process.cwd());
const outIndex = process.argv.indexOf('--out');
const outputPath = path.resolve(
  cleanArg(outIndex >= 0 ? process.argv[outIndex + 1] : '')
    || path.join(inputDir, '机种清单_合并.xlsx')
);

if (!fs.existsSync(inputDir) || !fs.statSync(inputDir).isDirectory()) {
  console.error(`目录不存在: ${inputDir}`);
  process.exit(1);
}

try {
  const result = await mergeReverseBom(inputDir, outputPath);
  console.log(`处理 PDF: ${result.files} 个`);
  console.log(`机种清单: ${result.rows} 条`);
  console.log(`输出: ${result.outputPath}`);
} catch (error) {
  console.error(`合并失败: ${error.message}`);
  process.exit(1);
}
