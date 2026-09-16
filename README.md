# ECN PDF → Excel 转换工具

将 ECN（设计变更通知，PDF 格式）逐个转换为：
- `xxx.txt`：PDF 原始文本解析结果
- `xxx_out.xlsx`：包含 `主表`、`子表` 两个工作表的 Excel 文件

## 环境要求（一次性安装）
- Node.js 18 及以上（推荐 20+）
- 使用 `pdf-parse` 解析 PDF，使用 `xlsx` 生成 Excel

## 数据格式依据
- `ec_pdf.txt` 是本工具对应的 ABAP 开发源码，记录 ECN 报表的字段、标签和明细布局。
- PDF 的字段解析规则依据该 ABAP 报表输出格式编写；如果 ABAP 报表布局发生变化，需要同步调整 `ecn_pdf_to_excel.mjs` 中的解析标签和明细规则。
- `ec_pdf.txt` 仅作为格式和字段参考，不是 Node.js 运行时依赖。

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

## 合并 ECN Excel

将批量转换生成的多个 `*_out_数字.xlsx` 文件合并为一个 Excel。脚本会读取每个文件的 `主表` 和 `子表1`，输出工作表 `主表`、`子表`。

```bash
# 默认读取当前目录下的 output，输出 output/miss_ec.xlsx
node merge_excel.mjs

# 指定输入目录
node merge_excel.mjs "/path/to/output"

# 指定输入目录和输出文件
node merge_excel.mjs "/path/to/output" --out "/path/to/output/miss_ec.xlsx"
```

输入目录中至少需要有一个符合 `*_out_数字.xlsx` 命名规则的文件，并且文件包含对应的工作表。

## 逆展開报告生成机种清单

针对 `ZS0PR012D` 最终逆展開报告，生成只包含以下四列的 Excel：
`親品目コード`、`親品目テキスト`、`直上品目コード`、`直上品目テキスト`。

```bash
node reverse_bom_to_excel.mjs "/path/to/ZS0PR012D.pdf"
node reverse_bom_to_excel.mjs "/path/to/ZS0PR012D.pdf" --out "./output"

# 或使用 npm 脚本
npm run reverse-bom -- "/path/to/ZS0PR012D.pdf"
```

输出文件名为 `原PDF文件名_机种清单.xlsx`，工作表名称为 `机种清单`。

多个逆展開 PDF 合并为一个 Excel：

```bash
node merge_reverse_bom.mjs "/path/to/pdf文件夹"
node merge_reverse_bom.mjs "/path/to/pdf文件夹" --out "./output/机种清单.xlsx"

# 或使用 npm 脚本
npm run merge-reverse-bom -- "/path/to/pdf文件夹"
```

脚本会读取输入目录内全部 PDF，合并输出一个 `机种清单` 工作表，并只保留上述四列。这里的输入目录是 PDF 所在目录，例如 `/path/to/pdf文件夹`。

## 逆展開物料级联回填

以 `stpo_mast/output/re_sc.xlsx` 中的 `親品目` 作为固定清单，再通过 `stpo_ec*.XLSX` 和 `mast_ec*.XLSX` 逐级推导完成品物料。这里的输入目录是 Excel 源文件所在目录，不是 PDF 目录。
级联关系按目录中实际存在的文件组自动判断，不限制层数。例如：

```text
stpo_ec* → mast_ec* → stpo_ec* → mast_ec* → ...
```

脚本会自动扫描 `stpo_ec*.XLSX`，按相同后缀匹配 `mast_ec*.XLSX`，并按文件编号顺序级联到最后一组。每一级使用 `stpo` 的 `组件` 匹配当前物料，再通过 `物料单` 关联同级 `mast` 的 `物料单`，取得下一层物料。没有下游 BOM 的物料作为最终完成品物料。

```bash
node fill_reverse_lookup.mjs "./stpo_mast"
node fill_reverse_lookup.mjs "./stpo_mast" --out "./output/re_lookup.XLSX"
```

脚本读取 `output/re_sc.xlsx`，默认输出到 `output/re_lookup.XLSX`；也可以通过 `--out` 指定输出文件，并生成两个工作表：

- `lookup`：每个 `親品目` 对应的完成品物料汇总在同一行。
- `list`：每个完成品物料单独一行，列为 `設変番号`、`完成品物料`、`親品目`。

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
| `merge_excel.mjs` | 合并多个 ECN Excel → 一个 Excel |
| `reverse_bom_to_excel.mjs` | 逆展開报告 → 机种清单 |
| `merge_reverse_bom.mjs` | 合并多个逆展開报告 → 一个机种清单 |
| `fill_reverse_lookup.mjs` | 按 stpo/mast 级联回填完成品物料 |
| `ec_pdf.txt` | 对应的 ABAP 开发源码和报表格式参考 |
| `package.json` | 依赖声明 |
| `README.md` | 本说明 |
