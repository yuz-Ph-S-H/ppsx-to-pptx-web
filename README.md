# PPSX → PPTX 批量转换网页

一个纯前端、本地处理的 PPSX → PPTX 批量转换工具。

## 功能

- 选择一个本地文件夹。
- 扫描该文件夹最外层的全部 `.ppsx` 文件。
- 在浏览器本地转换，不上传文件。
- 生成同名 `.pptx` 并直接写回原文件夹。
- 保留原 `.ppsx`。
- 若同名 `.pptx` 已存在，会在开始前提示，确认后覆盖。
- 显示文件列表、单文件状态、总体进度和完成/失败提示。

## 本地运行（Windows）

### 方法 1：直接双击

1. 解压本项目。
2. 双击 `start.bat`。
3. 浏览器会打开 `http://localhost:8000`。
4. 推荐使用最新版 Microsoft Edge 或 Google Chrome。
5. 点击“选择文件夹”，选择包含 PPSX 的文件夹。
6. 点击“开始转换”。

`start.bat` 需要电脑已经安装 Python 3（`py` 或 `python` 命令可用）。

### 方法 2：命令行

在项目目录执行：

```bash
python -m http.server 8000 --bind 127.0.0.1
```

然后打开：

```text
http://localhost:8000
```

## 为什么不能直接双击 index.html

该网页需要浏览器的 File System Access API 来获得用户主动授权后的文件夹读写权限。文件夹选择与写回文件需要安全上下文；本地测试时使用 `http://localhost`。

## 转换原理

`.ppsx` 和 `.pptx` 都属于 Office Open XML 演示文稿包。工具读取 ZIP 包中的 `[Content_Types].xml`，把 PowerPoint Show 主文档类型：

```text
application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml
```

替换为 PowerPoint Presentation 主文档类型：

```text
application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml
```

随后重新打包，并以 `.pptx` 文件名写回原目录。

## 浏览器建议

优先使用：

- Microsoft Edge（最新版）
- Google Chrome（最新版）

其他浏览器对 `showDirectoryPicker()` 的支持可能不同。

## GitHub Pages

该项目是纯静态网页，后续可以直接部署到 GitHub Pages。公网部署需要 HTTPS；GitHub Pages 默认提供 HTTPS。

## 注意事项

- 当前只扫描所选文件夹最外层，不递归子文件夹。
- 只支持 `.ppsx`，不处理 `.ppsm` 等含宏格式。
- 修改 OOXML 包会使原有数字签名失效（如果文件存在数字签名）。
- 对超大演示文稿，浏览器转换时会占用与文件大小相关的内存。

## 第三方库

项目内置 JSZip 3.10.1，用于在浏览器内读取和重新生成 OOXML ZIP 包。JSZip 使用 MIT / GPLv3 双许可证，许可证文本见 `vendor/JSZip-LICENSE.markdown`。
