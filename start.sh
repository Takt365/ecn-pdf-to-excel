#!/usr/bin/env bash
# Mac / Linux 一键启动脚本
set -e
cd "$(dirname "$0")"

if [ ! -d "node_modules" ]; then
  echo "📦 首次运行，安装依赖..."
  npm install
fi

echo ""
echo "1) 单个PDF转换"
echo "2) 批量转换整个文件夹"
read -rp "请选择 (1/2): " CHOICE

case "$CHOICE" in
  1)
    read -rp "请输入PDF文件路径: " PDF
    node ecn_pdf_to_excel.js "$PDF"
    ;;
  2)
    read -rp "请输入包含PDF的文件夹路径: " DIR
    node batch.js "$DIR"
    ;;
  *)
    echo "无效选择"
    exit 1
    ;;
esac
