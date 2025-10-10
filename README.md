# ppt2html

使用 Go 编写的命令行工具，可以将 PPT/PPTX 演示文稿转换为包含原始布局和样式信息的静态 HTML 页面。

## 功能特点

- 解析幻灯片尺寸、背景、文本形状和图片，尽力还原原始位置、大小与样式。
- 输出包含 `index.html` 和对应资源文件（图片等）的目录，可直接在浏览器中打开。
- 纯 Go 实现，无需依赖本地安装 Office 或 LibreOffice。

## 安装

```bash
go build ./cmd/ppt2html
```

## 使用示例

```bash
ppt2html -input sample.pptx -output output-dir
```

生成后的 `output-dir/index.html` 即为 HTML 预览文件，`output-dir/assets` 中存放演示文稿中引用的图片等资源。

## 注意事项

- 该工具目前主要支持 PPTX（Open XML）文件。对于旧版二进制 `.ppt` 文件，可先使用 Office/LibreOffice 转换为 `.pptx`。
- 某些复杂动画、嵌入对象或特殊样式可能无法完全还原，但核心布局和基础样式会被保留。
- 如遇无法解析的特殊内容，可扩展 `internal/pptx` 包中的解析逻辑，或结合大模型生成额外的 HTML/CSS 片段进行增强。

