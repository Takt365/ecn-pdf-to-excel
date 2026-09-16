#!/usr/bin/env node
import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';

function cleanArg(value) {
  return value?.replace(/^['"]|['"]$/g, '');
}

const rootDir = path.resolve(cleanArg(process.argv[2]) || 'stpo_mast');
const outIndex = process.argv.indexOf('--out');
const inputLookupPath = path.join(rootDir, 'output', 're_sc.xlsx');
const outputLookupPath = path.resolve(
  cleanArg(outIndex >= 0 ? process.argv[outIndex + 1] : '')
    || path.join(rootDir, 'output', 're_lookup.XLSX')
);

if (!fs.existsSync(rootDir) || !fs.statSync(rootDir).isDirectory()) {
  console.error(`输入目录不存在: ${rootDir}`);
  process.exit(1);
}

ensureOutputDir(outputLookupPath);

if (!fs.existsSync(inputLookupPath)) {
  console.error(`输入文件不存在: ${inputLookupPath}`);
  console.error('请先准备 output/re_sc.xlsx，且至少包含一列 親品目。');
  process.exit(1);
}

function readRows(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function text(value) {
  return String(value ?? '').trim();
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

const levelFiles = fs.readdirSync(rootDir)
  .map(name => {
    const match = name.match(/^stpo_ec(.+)\.xlsx$/i);
    return match ? { suffix: match[1], stpoName: name } : null;
  })
  .filter(level => level && fs.existsSync(path.join(rootDir, `mast_ec${level.suffix}.XLSX`)))
  .sort((left, right) => left.suffix.localeCompare(right.suffix, undefined, { numeric: true }));

if (levelFiles.length === 0) {
  console.error(`未找到成对的 stpo_ec*.XLSX 和 mast_ec*.XLSX: ${rootDir}`);
  process.exit(1);
}

const lookupRows = readRows(inputLookupPath).map(row => ({
  設変番号: text(row['設変番号']),
  親品目: text(row['親品目']),
  完成品物料: ''
})).filter(row => row['親品目']);
const levels = [];
for (const level of levelFiles) {
  const stpoRows = readRows(path.join(rootDir, level.stpoName));
  const mastRows = readRows(path.join(rootDir, `mast_ec${level.suffix}.XLSX`));
  const mastByBom = new Map();
  for (const row of mastRows) {
    const bom = text(row['物料单']);
    const material = text(row['物料']);
    if (!bom || !material) continue;
    if (!mastByBom.has(bom)) mastByBom.set(bom, new Set());
    mastByBom.get(bom).add(material);
  }
  const nextByComponent = new Map();
  for (const row of stpoRows) {
    const component = text(row['组件']);
    const bom = text(row['物料单']);
    if (!component || !bom || !mastByBom.has(bom)) continue;
    if (!nextByComponent.has(component)) nextByComponent.set(component, new Set());
    for (const material of mastByBom.get(bom)) nextByComponent.get(component).add(material);
  }
  levels.push(nextByComponent);
}

for (const row of lookupRows) {
  const parent = text(row['親品目']);
  let currentMaterials = parent ? new Set([parent]) : new Set();
  const finalMaterials = new Set();
  for (const nextByComponent of levels) {
    const nextMaterials = new Set();
    for (const material of currentMaterials) {
      const next = nextByComponent.get(material);
      if (!next || next.size === 0) {
        nextMaterials.add(material);
        continue;
      }
      for (const nextMaterial of next) nextMaterials.add(nextMaterial);
    }
    currentMaterials = nextMaterials;
  }
  for (const material of currentMaterials) finalMaterials.add(material);
  row['完成品物料'] = [...finalMaterials].join(', ');
}

const lookupHeaders = ['設変番号', '親品目', '完成品物料'];
const outputSheet = XLSX.utils.json_to_sheet(lookupRows, { header: lookupHeaders });
const listRows = [];
for (const row of lookupRows) {
  const changeNumber = text(row['設変番号']);
  const parent = text(row['親品目']);
  const materials = text(row['完成品物料'])
    .split(',')
    .map(value => value.trim())
    .filter(Boolean);
  for (const material of materials) {
    listRows.push({ '設変番号': changeNumber, '完成品物料': material, '親品目': parent });
  }
}
const listSheet = XLSX.utils.json_to_sheet(listRows, {
  header: ['設変番号', '完成品物料', '親品目']
});
const lookupWorkbook = XLSX.utils.book_new();
lookupWorkbook.Sheets = {};
lookupWorkbook.SheetNames = [];
XLSX.utils.book_append_sheet(lookupWorkbook, outputSheet, 'lookup');
XLSX.utils.book_append_sheet(lookupWorkbook, listSheet, 'list');
XLSX.writeFile(lookupWorkbook, outputLookupPath);

const matched = lookupRows.filter(row => text(row['完成品物料'])).length;
console.log(`親品目: ${lookupRows.length} 条`);
console.log(`已回填: ${matched} 条`);
console.log(`未匹配: ${lookupRows.length - matched} 条`);
console.log(`list 明细: ${listRows.length} 条`);
console.log(`级联层数: ${levels.length} 组`);
console.log(`输出: ${outputLookupPath}`);
