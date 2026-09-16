#!/usr/bin/env node
/**
 * 批量处理文件夹内所有 PDF
 * 用法: node batch.js <文件夹路径>
 */

import { execSync } from "child_process";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// 取第一个参数作为文件夹路径
const dir = process.argv[2];

if (!dir) {
  console.log("用法: node batch.js <文件夹路径> [--out <输出文件夹>]");
  console.log('例:  node batch.js "G:\\AppDevelop\\VS2026\\基础数据\\ec_pdf"');
  process.exit(1);
}

// 去掉首尾引号（Windows 有时会带）
const cleanDir = dir.replace(/^["']|["']$/g, "");
const outIndex = process.argv.indexOf("--out");
const outDir = outIndex >= 0 && process.argv[outIndex + 1]
  ? process.argv[outIndex + 1].replace(/^["']|["']$/g, "")
  : path.join(cleanDir, "output");

if (!fs.existsSync(cleanDir)) {
  console.error("❌ 文件夹不存在:", cleanDir);
  process.exit(1);
}

const files = fs.readdirSync(cleanDir).filter((f) =>
  f.toLowerCase().endsWith(".pdf")
);

if (files.length === 0) {
  console.log("⚠️ 该文件夹下没有找到 PDF 文件");
  process.exit(0);
}

console.log(`📁 找到 ${files.length} 个 PDF 文件:\n`);

let success = 0;
let fail = 0;

for (const f of files) {
  const fullPath = path.join(cleanDir, f);
  console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
  console.log(`📄 处理: ${f}`);
  console.log(`   路径: ${fullPath}`);
  try {
    execSync(`node "${path.join(__dirname, "ecn_pdf_to_excel.mjs")}" "${fullPath}" --out "${outDir}"`, {
      stdio: "inherit",
    });
    success++;
  } catch (e) {
    console.error(`❌ 失败: ${f}`);
    fail++;
  }
}

console.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
console.log(`🏁 完成! 成功: ${success}, 失败: ${fail}, 总计: ${files.length}`);