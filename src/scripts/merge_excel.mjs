import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';
import { fileURLToPath } from 'url';

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

function readRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
}

function mergeExcel(inputDir, outputPath) {
  const files = fs.readdirSync(inputDir)
    .filter(name => /^.+_out_\d+\.xlsx$/i.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  if (files.length === 0) {
    throw new Error(`目录中没有找到 *_out_行数.xlsx 文件: ${inputDir}`);
  }

  let mainHeader = null;
  let subHeader = null;
  const mainRows = [];
  const subRows = [];

  for (const file of files) {
    const filePath = path.join(inputDir, file);
    const workbook = XLSX.readFile(filePath);
    const mainSheetName = workbook.SheetNames.find(name => name === '主表');
    const subSheetName = workbook.SheetNames.find(name => /^子表\d+$/u.test(name));

    if (!mainSheetName) continue;

    const mainData = readRows(workbook.Sheets[mainSheetName]);
    if (mainData.length > 0) {
      mainHeader ??= mainData[0];
      mainRows.push(...mainData.slice(1));
    }

    if (subSheetName) {
      const subData = readRows(workbook.Sheets[subSheetName]);
      if (subData.length > 0) {
        subHeader ??= subData[0];
        subRows.push(...subData.slice(1));
      }
    }
  }

  if (!mainHeader) throw new Error('没有读取到主表表头');
  if (!subHeader) throw new Error('没有读取到子表表头');

  const outputWorkbook = XLSX.utils.book_new();
  const mainSheet = XLSX.utils.aoa_to_sheet([mainHeader, ...mainRows]);
  const subSheet = XLSX.utils.aoa_to_sheet([subHeader, ...subRows]);
  XLSX.utils.book_append_sheet(outputWorkbook, mainSheet, '主表');
  XLSX.utils.book_append_sheet(outputWorkbook, subSheet, '子表');
  ensureOutputDir(outputPath);
  XLSX.writeFile(outputWorkbook, outputPath);

  return { files: files.length, mainRows: mainRows.length, subRows: subRows.length, outputPath };
}

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const resourcesDataDir = path.join(scriptDir, '..', 'resources', 'data');
const inputDir = path.resolve(cleanArg(process.argv[2]) || path.join(resourcesDataDir, 'ec_pdf', 'output'));
const outIndex = process.argv.indexOf('--out');
const outputPath = path.resolve(
  cleanArg(outIndex >= 0 ? process.argv[outIndex + 1] : '')
    || path.join(inputDir, 'miss_ec.xlsx')
);

if (!fs.existsSync(inputDir)) {
  console.error(`目录不存在: ${inputDir}`);
  process.exit(1);
}

try {
  const result = mergeExcel(inputDir, outputPath);
  console.log(`合并文件: ${result.files}`);
  console.log(`主表数据: ${result.mainRows} 条`);
  console.log(`子表数据: ${result.subRows} 条`);
  console.log(`输出: ${result.outputPath}`);
} catch (error) {
  console.error(`合并失败: ${error.message}`);
  process.exit(1);
}
