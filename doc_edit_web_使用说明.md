# doc_edit_web 使用说明

`doc_edit_web.py` 用于将原始文件（PNG/SVG/PDF/XLSX）导入到一个双面板网页：

- 左侧：原始文件导入与预览
- 右侧：OCR/解析后的表格与文本可视化编辑

编辑完成后可导出为 Excel（.xlsx）。

## 依赖

- Python 3.10+
- openpyxl（用于读写 Excel）
- PaddleOCR（用于 OCR 识别图片/文档截图）
- 可选依赖（仅当导入 PDF/SVG 时需要）
  - PDF 转 PNG：PyMuPDF（`fitz`）或 `pdf2image`
  - SVG 转 PNG：`cairosvg`

## 启动方式

在本目录打开 PowerShell：

### 1) 启动网页服务（推荐：从页面导入文件）

```powershell
python -u doc_edit_web.py --port 8012
```

打开浏览器访问：

- `http://127.0.0.1:8012/`

### 2) 可选：启动时预加载一个文件（兼容旧用法）

```powershell
python -u doc_edit_web.py --input 你的文件.xlsx --sheet Sheet1 --table-range A1:H200 --text-cells J1:J10 --port 8012
```

说明：

- `--input` 支持：`.xlsx/.xlsm/.png/.jpg/.jpeg/.bmp/.tif/.tiff/.pdf/.svg`
- 如果不填 `--input`，启动后在页面左侧导入即可
- Excel 相关参数（仅 Excel 导入生效）
  - `--sheet`：工作表名称；不填默认 active sheet
  - `--table-range`：表格区域（A1 样式，第一行视为表头，下面为数据区）
  - `--text-cells`：文本单元格地址（A1 样式，支持 `J1:J10` 或 `J1,J2,J3`）

## 页面使用步骤（双面板）

### 1) 左侧导入文件

- 点击“选择文件”或将文件拖拽到导入区域
- 支持：PNG/SVG/PDF/XLSX（页面里也允许选择 XLSM）
- 点击“导入并解析”

### 2) 右侧交互编辑

- 表格编辑：双击单元格直接修改
- 表格增删：点击“添加行”/“删除选中行”
- 查询/过滤：每列表头下方输入框可过滤该列
- 文本编辑：右侧文本框直接编辑
- 校验：导出前必须通过校验（必填、数字类型、范围等），错误会显示在“数据校验”
- 历史：页面右侧会持续记录操作历史

## 导出规则（格式保持）

### Excel 输入（格式尽量保持）

- 在原始 Excel 的基础上写回数据与文本
- 单元格样式（字体/边框/填充/对齐/格式等）会尽量保留
- 若新增行超过原表格数据区，会插入新行并复制“模板行”的样式
- 若你在页面导入 Excel，原始文件会保存到本目录的 `uploads/` 下，导出基于该文件进行写回

### PNG/SVG/PDF 输入（OCR 重建表格）

- 生成一个新的 `.xlsx` 文件
- 表格第一行为表头
- 文本内容会追加写在表格下方
- SVG/PDF 会先转成 PNG 再 OCR（需要安装可选依赖）

## 输出文件与缓存目录

- 导出的 Excel：浏览器直接下载
- 修改历史：导出时会在脚本目录生成 `*_history_YYYYmmdd_HHMMSS.json`
- 上传缓存：页面导入的原始文件与（必要时）转换后的 PNG 会保存到 `uploads/`

## 常见问题

### 1) 导出提示“存在校验错误”

右侧“数据校验”会列出错误位置（行/列）。修正后再导出。

### 2) Excel 的 table-range 怎么填

示例：表头在第 1 行，数据到第 200 行，列从 A 到 H（Excel 导入时使用）：

- `--table-range A1:H200`

如果你的表格不是从 A1 开始，也可以是：

- `--table-range B3:K80`

### 3) PDF/SVG 导入时报错“需要安装 xxx”

这是因为 PDF/SVG 需要先转换为 PNG：

- PDF：安装 PyMuPDF（`pip install PyMuPDF`）或 `pdf2image`
- SVG：安装 `cairosvg`
