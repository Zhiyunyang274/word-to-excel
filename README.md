# 📄 Word 转 Excel 在线工具

一个免费、纯前端的 Word 文档转 Excel 表格工具，支持在线使用，无需注册，文件本地处理，数据安全有保障。

## ✨ 功能特性

- **本地处理** — 文件完全在浏览器本地转换，不上传任何服务器，隐私安全
- **无需注册** — 打开即用，零门槛
- **表格智能识别** — 自动提取 Word 中的所有表格，每张表格生成独立的 Sheet
- **纯文本兜底** — 若文档中无表格，自动将正文按行提取为带序号的 Excel
- **自适应列宽** — 根据内容自动计算合适的列宽
- **拖拽上传** — 支持拖拽文件到上传区域
- **手机端友好** — 移动端完全适配，随时随地使用

## 🚀 在线体验

> 将 `index.html` 部署到任意静态托管服务（如 GitHub Pages、Vercel、Netlify）即可使用。

## 🛠️ 技术栈

| 依赖 | 用途 |
|------|------|
| [mammoth.js](https://github.com/mwilliamson/mammoth.js) | 解析 `.docx` 文件，转换为 HTML |
| [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) | 生成并导出 `.xlsx` 文件 |

所有依赖均通过 CDN 加载，无需安装任何包。

## 📦 项目结构

```
word-to-excel/
├── index.html   # 页面结构与样式
├── app.js       # 核心转换逻辑
└── README.md
```

## 🖥️ 本地运行

直接用浏览器打开 `index.html` 即可，无需任何构建步骤：

```bash
# 或者用 VS Code Live Server、Python 起一个本地服务
python -m http.server 8080
```

然后访问 `http://localhost:8080`。

## 📋 使用方法

1. 点击上传区域（或拖拽文件）选择 `.docx` 文件
2. 点击「开始转换」按钮
3. 转换完成后，`.xlsx` 文件自动下载

## ⚠️ 注意事项

- 目前仅支持 `.docx` 格式（不支持旧版 `.doc`）
- 复杂的合并单元格、嵌套表格可能无法完美还原
- 文件大小受浏览器内存限制，建议单文件不超过 50MB

## 📄 License

[MIT](LICENSE)
