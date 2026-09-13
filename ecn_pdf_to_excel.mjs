import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';
import { PDFParse } from 'pdf-parse';

const MAIN_COLUMNS = [
  '設変番号', '機種', 'タイトル', 'ステータス', '発行日', '担当', '依頼元', '設変会議',
  'PP番号', '技術連絡書№', '実施', '変更理由（主）', '変更理由（従）', '安全規格',
  '進行状況', '機番管理', '客先承認', 'Sマニュアル訂正', '取設訂正', 'カタログ訂正',
  '作業標準訂正', '技術情報発行', 'コスト変動', '単位コスト', '金型改造費', '関連図面', '設変記事'
];

const SUB_COLUMNS = [
  '行号', '設変番号', '完成品', '親品目', '旧品目', '旧品目テキスト', '旧数量', '旧取付位置',
  '新品目', '新品目テキスト', '新数量', '新取付位置', '明細番号', '互換性', 'セカンド区分',
  '手配指示', '旧部品処理', '有効日付'
];

const LABELS = {
  '機種': ['機種'],
  '発行日': ['設変会議日', '発行日'], '担当': ['担当'], '依頼元': ['依頼元'],
  '設変会議': ['設変会議'], 'PP番号': ['ＰＰ番号', 'PP番号', 'P.P.NO.'],
  '技術連絡書№': ['技術連絡書№', 'ENG.INFO.NO.'], '実施': ['実施'],
  '変更理由（主）': ['変更理由（主）', 'REASON(MAIN)'],
  '変更理由（従）': ['変更理由（従）', 'REASON(SUB)'],
  '安全規格': ['安全規格', 'SAFETY STD'], '進行状況': ['進行状況', 'COORD.TO MERCHANT'],
  '機番管理': ['機番管理', 'CONTROL LEVEL'], '客先承認': ['客先承認', 'CUSTOMER"S APPRV'],
  'Sマニュアル訂正': ['Ｓマニュアル訂正', 'Sマニュアル訂正', 'CHG.IN SERV.MANUAL'],
  '取設訂正': ['取説訂正', 'CHG.IN USR. MANUAL'], 'カタログ訂正': ['カタログ訂正', 'CHG.IN BROCHURE'],
  '作業標準訂正': ['作業標準訂正', 'CHG.IN MFTR.MANUAL'],
  '技術情報発行': ['技術情報発行', 'ENGINEERING INFO.'], 'コスト変動': ['コスト変動', 'COST'],
  '単位コスト': ['単位コスト', 'UNIT COST'], '金型改造費': ['金型改造費', 'COST FOR DIE'],
  '関連図面': ['関連図面', 'RELATED DRAWINGS']
};

const MAIN_JA_LABELS = [
  '機種', '変更番号', 'タイトル', 'ステータス', '発行日', '設変会議日', '担当', '依頼元',
  '設変会議', '同期設変', '関連設変', 'ＰＰ番号', '技術連絡書№', '実施', '変更理由（主）',
  '変更理由（従）', '安全規格', '進行状況', '機番管理', '客先承認', 'Ｓマニュアル訂正',
  '取説訂正', 'カタログ訂正', '作業標準訂正', '技術情報発行', 'コスト変動', '単位コスト',
  '金型改造費', '関連図面', '設変記事'
];

function clean(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function labelValue(text, aliases) {
  const label = aliases.map(escapeRegExp).sort((a, b) => b.length - a.length).join('|');
  const otherLabels = [
    ...Object.values(LABELS).flat(), ...MAIN_JA_LABELS
  ].map(escapeRegExp).join('|');
  const pattern = new RegExp(
    `(?:${label})[ \\t]*[:：][ \\t]*([^\\n]*?)(?=[ \\t]+(?:${otherLabels})[ \\t]*[:：]|[ \\t]+(?:DATE|STATUS|TITLE|PRODUCTS|E\\.C\\.N\\.№|PERSON IN CHARGE|REQUESTING SECTION|E\\.C\\.MTG\\.|SAFETY STD|COORD\\.TO MERCHANT|CONTROL LEVEL|CUSTOMER["']S APPRV|RELATED DRAWINGS|CHG\\.IN|ENGINEERING INFO\\.|COST FOR DIE|UNIT COST)\\b|$)`,
    'im'
  );
  return clean(text.match(pattern)?.[1]);
}

function parseMain(text) {
  const firstPage = text.split(/\f|\nREPORT-ID\s*:\s*\S+\s+PAGE\s*:\s*2\b|\n--\s*1\s+of\s+/i)[0];
  const main = Object.fromEntries(MAIN_COLUMNS.map(name => [name, '']));
  main['設変番号'] = clean(firstPage.match(/(?:変更番号|E\.C\.N\.№)\s*[:：]?\s*([^\s]+)/i)?.[1]);
  main['タイトル'] = labelValue(firstPage, ['タイトル']);
  main['ステータス'] = clean(firstPage.match(/ステータス\s*[:：]\s*([^\)\n]+)/i)?.[1]);
  for (const [name, aliases] of Object.entries(LABELS)) main[name] = labelValue(firstPage, aliases);

  main['設変記事'] = parseMemo(firstPage);
  return main;
}

function parseMemo(firstPage) {
  const rightLabels = [
    '実施', '変更理由（主）', '変更理由（従）', '安全規格', '進行状況', '機番管理',
    '客先承認', 'Ｓマニュアル訂正', '取説訂正', 'カタログ訂正', '作業標準訂正',
    '技術情報発行', 'コスト変動', '単位コスト', '金型改造費', '関連図面',
    'ENFORCEMENT', 'REASON\\(MAIN\\)', 'REASON\\(SUB\\)', 'SAFETY STD',
    'COORD\\.TO MERCHANT', 'CONTROL LEVEL', "CUSTOMER[\"']S APPRV",
    'CHG\\.IN', 'ENGINEERING INFO\\.', 'COST', 'UNIT COST', 'COST FOR DIE',
    'RELATED DRAWINGS'
  ].join('|');
  const lines = firstPage.split(/\r?\n/);
  const result = [];
  let collecting = false;

  for (const rawLine of lines) {
    let line = rawLine;
    if (!collecting) {
      const contents = line.match(/CONTENTS\s*/i);
      if (!contents) continue;
      collecting = true;
      line = line.slice(contents.index + contents[0].length);
    }

    if (new RegExp(`^\\s*(?:${rightLabels})[ \\t]*[:：]`, 'i').test(line)) continue;
    line = line.replace(new RegExp(`[ \\t]+(?:${rightLabels})[ \\t]*[:：]?.*$`, 'i'), '');
    line = line.replace(/^\s*(?:ENFORCEMENT|REASON\(MAIN\)|REASON\(SUB\)|SAFETY STD|COORD\.TO MERCHANT|CONTROL LEVEL|CUSTOMER["']S APPRV|RELATED DRAWINGS|CHG\.IN.*|ENGINEERING INFO\.|COST|UNIT COST|COST FOR DIE)\s*$/i, '');
    line = line.trim();
    if (line) result.push(line);
  }

  return result.join('\n').trim();
}

// ABAP material numbers are one or two letters followed by 7-9 digits,
// optionally ending in a revision letter. This excludes uppercase words in descriptions.
const ITEM_RE = /\b[A-Z]{1,2}\d{7,9}[A-Z]?\b/g;
const QTY_RE = /\b\d+\.\d{3}\b/g;

function parseSide(data, itemMatch, quantity) {
  if (!itemMatch || !quantity) return { code: '', name: '', quantity: '', position: '' };
  const codeEnd = itemMatch.index + itemMatch[0].length;
  const quantityAt = data.indexOf(quantity, codeEnd);
  return {
    code: itemMatch[0],
    name: clean(data.slice(codeEnd, quantityAt)),
    quantity,
    position: ''
  };
}

function parseDetailLine(line, current, mainNo) {
  const row = line.match(/^(\d+)\s*-\s*(\d+)\s+([\s\S]+)$/);
  if (!row) return null;
  const meta = row[3].match(/\s+(\d{4})\s+([A-D][*]?|-)\s+([1-9-])\s+([1-9-])\s+([1-9-])\s*$/);
  const data = meta ? row[3].slice(0, meta.index) : row[3];
  const items = [...data.matchAll(ITEM_RE)];
  const quantities = [...data.matchAll(QTY_RE)].map(match => match[0]);
  let oldSide = parseSide(data, items[0], quantities[0]);
  let newSide = parseSide(data, items[1], quantities[1]);
  if (items.length === 1) {
    // With one material, the ABAP report writes the detail flags only on the old side.
    if (!meta) {
      newSide = oldSide;
      oldSide = { code: '', name: '', quantity: '', position: '' };
    }
  }

  const potx = [meta?.[2] || '', meta?.[3] || '', meta?.[4] || '', meta?.[5] || ''];
  return {
    '行号': `${row[1]}-${row[2]}`,
    '設変番号': mainNo,
    '完成品': '',
    '親品目': current.parent,
    '旧品目': oldSide.code, '旧品目テキスト': oldSide.name, '旧数量': oldSide.quantity,
    '旧取付位置': oldSide.position, '新品目': newSide.code, '新品目テキスト': newSide.name,
    '新数量': newSide.quantity, '新取付位置': newSide.position, '明細番号': meta?.[1] || '',
    '互換性': potx[0], 'セカンド区分': potx[1], '手配指示': potx[2], '旧部品処理': potx[3],
    '有効日付': ''
  };
}

function parseSub(text, mainNo) {
  const rows = [];
  let current = { seq: '', parent: '' };
  let currentRow = null;
  const lines = text.replace(/<PARSED TEXT FOR PAGE: \d+ \/ \d+>/g, '').split(/\r?\n/);
  for (const raw of lines) {
    const line = clean(raw);
    const parent = line.match(/^親品目\s*[:：]\s*(\S+)/);
    if (parent) {
      current.parent = parent[1];
      continue;
    }
    const material = line.match(/^MATERIAL\s*\(\s*(\d+)\s*\)/i);
    if (material) {
      current.seq = material[1];
      continue;
    }
    if (/^\d+\s*-\s*\d+\s+/.test(line)) {
      const row = parseDetailLine(line, current, mainNo);
      if (row) {
        rows.push(row);
        currentRow = row;
      }
      continue;
    }
    const bothPositions = line.match(/^(\d{4})\s+(\d+\.\d{3})\s+(\S+)\s+(\d{4})\s+(\d+\.\d{3})\s+(\S+)$/);
    const onePosition = line.match(/^(\d{4})\s+(\d+\.\d{3})\s+(\S+)$/);
    if (currentRow && bothPositions) {
      currentRow['旧取付位置'] = appendPosition(currentRow['旧取付位置'], bothPositions[3]);
      currentRow['新取付位置'] = appendPosition(currentRow['新取付位置'], bothPositions[6]);
      continue;
    }
    if (currentRow && onePosition) {
      const target = currentRow['旧数量'] ? '旧取付位置' : '新取付位置';
      currentRow[target] = appendPosition(currentRow[target], onePosition[3]);
      continue;
    }
    if (/^(?:REPORT-ID|設計変更情報|ENGINEERING CHANGE|旧品目|COMPONENT|MATERIAL|\+)/i.test(line)) continue;
  }
  return rows;
}

function appendPosition(existing, value) {
  if (!value || existing.split(',').includes(value)) return existing;
  return existing ? `${existing},${value}` : value;
}

function exportExcel(main, subs, outputPath) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([main], { header: MAIN_COLUMNS }), '主表');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(subs, { header: SUB_COLUMNS }), `子表${subs.length}`);
  XLSX.writeFile(workbook, outputPath);
}

export async function convertPdf(pdfPath, outDir) {
  const base = path.basename(pdfPath, path.extname(pdfPath));
  fs.mkdirSync(outDir, { recursive: true });
  const parser = new PDFParse({ data: fs.readFileSync(pdfPath) });
  const parsed = await parser.getText();
  const txtPath = path.join(outDir, `${base}.txt`);
  fs.writeFileSync(txtPath, parsed.text, 'utf8');
  const main = parseMain(parsed.text);
  const subs = parseSub(parsed.text, main['設変番号'] || base);
  const excelPath = path.join(outDir, `${base}_out_${subs.length}.xlsx`);
  exportExcel(main, subs, excelPath);
  return { txtPath, excelPath, main, subs };
}

const pdfPath = process.argv[2]?.replace(/^['"]|['"]$/g, '');
if (!pdfPath || !fs.existsSync(pdfPath)) {
  console.error('用法: node ecn_pdf_to_excel.mjs <PDF路径> [--out <输出目录>]');
  process.exit(1);
}
const outIndex = process.argv.indexOf('--out');
const outDir = outIndex >= 0 && process.argv[outIndex + 1]
  ? path.resolve(process.argv[outIndex + 1])
  : path.join(path.dirname(path.resolve(pdfPath)), 'output');
const result = await convertPdf(pdfPath, outDir);
console.log(`TXT: ${result.txtPath}`);
console.log(`Excel: ${result.excelPath}`);
console.log(`主表 1 条，子表 ${result.subs.length} 条`);
