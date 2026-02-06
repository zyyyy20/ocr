# ocr

本目录用于在 Windows 上快速验证 PaddleOCR 的本地 OCR 识别与可视化输出。

## 环境版本（当前机器实测）

- Python: 3.10.7
- PaddlePaddle（paddle）: 3.2.2
- PaddleOCR（paddleocr）: 3.4.0

安装步骤见：`环境安装.md`

## 目录文件说明

- `run_ocr_local.py`
  - 功能：输出 Python / paddle / paddleocr 版本，并对 `IMAGE_PATH` 指定图片做 OCR，打印每行的 `文本 + 置信度`
  - 入口：直接运行脚本
- `run_ocr_visualize.py`
  - 功能：对 `IMAGE_PATH` 指定图片做 OCR，并将检测框 + `文字 + 置信度` 画到图片上
  - 输出：`ocr_visualization.png`（保存到本目录）
- `doc_edit_web.py`
  - 功能：双面板网页导入 PNG/SVG/PDF/XLSX，经 OCR/解析后可交互编辑并导出 Excel
  - 使用方式：`doc_edit_web_使用说明.md`
- `ocr_visualization.png`
  - 功能：`run_ocr_visualize.py` 的默认输出图片（可覆盖生成）

## 如何运行

在本目录打开 PowerShell：

### 1) 纯文本 OCR（打印识别结果）

```powershell
python -u run_ocr_local.py
```

修改识别图片：编辑 `run_ocr_local.py` 顶部的 `IMAGE_PATH`。

### 2) OCR 可视化（生成带框/文字/置信度的图片）

```powershell
python -u run_ocr_visualize.py
```

修改识别图片：编辑 `run_ocr_visualize.py` 顶部的 `IMAGE_PATH`。

输出路径：脚本内 `OUTPUT_PATH`，默认写到本目录下的 `ocr_visualization.png`。

## 兼容性说明

- 本目录脚本会设置一些 Paddle 相关环境变量以尽量规避 CPU(oneDNN) 推理兼容问题。
- 如遇到报错包含 `ConvertPirAttribute2RuntimeAttribute` / `onednn_instruction.cc`，通常与 `paddlepaddle==3.3.0` 的 CPU(oneDNN) 推理相关；当前环境已使用 `paddlepaddle==3.2.2` 通过验证。

