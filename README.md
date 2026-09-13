# ECN PDF → Excel 转换工具

将 ECN（设计变更通知，PDF 格式）逐个转换为：
- `xxx.txt`：PDF 原始文本解析结果
- `xxx_out.xlsx`：包含 `主表`、`子表` 两个工作表的 Excel 文件

## 环境要求（一次性安装）
- Node.js 18 及以上（推荐 20+）
- 使用 `pdf-parse` 解析 PDF，使用 `xlsx` 生成 Excel

```bash
# Windows: 下载安装 Node.js https://nodejs.org
# Mac:      brew install node
# Linux:    sudo apt install nodejs npm
```

## 快速开始

```bash
# 1. 解压后进入目录
cd ecn-pdf-to-excel

# 2. 安装依赖（只需一次）
npm install

# 3. 单个 PDF 转换
node ecn_pdf_to_excel.mjs "/path/to/GA1674.pdf"
#   → 在同目录生成两个 Excel

# 3b. 指定输出目录
node ecn_pdf_to_excel.mjs "/path/to/GA1674.pdf" --out "./output"

# 4. 批量转换整个文件夹
node batch.js "/path/to/pdf文件夹"
node batch.js "/path/to/pdf文件夹" --out "./output"
```

也可用 npm 脚本：
```bash
npm start -- "/path/to/GA1674.pdf"
npm run batch -- "/path/to/pdf文件夹"
```

## 输出示例
```
✅ 使用 pdftotext 解析
  ✅ GA1674.txt
  ✅ GA1674_out.xlsx
🎉 完成！
```

## 常见问题
- **报 `xlsx` 或 `pdf-parse` 找不到** → 确认在目录下执行过 `npm install`
- **解析字段为空** → 不同 ECN 模板字段名可能不同，按 `ecn_pdf_to_excel.mjs` 里的 `pickLine()` 调整标签即可
- **中文/日文乱码** → 确保终端与系统区域设置支持 UTF-8

## 文件说明
| 文件 | 作用 |
|------|------|
| `ecn_pdf_to_excel.mjs` | 主脚本：单个 PDF → 两个 Excel |
| `batch.js` | 批量脚本：遍历文件夹 |
| `package.json` | 依赖声明 |
| `README.md` | 本说明 |
