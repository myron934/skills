---
name: "pptx-to-md"
description: "Converts PPTX files to Markdown with image extraction and OCR. Invoke when user wants to convert/extract PPT/PPTX content to Markdown or needs to extract text from PowerPoint files."
---

# PPTX to Markdown Skill

将PPTX文件解析为Markdown格式，支持图片提取和OCR文字识别。

## 功能特点

- 提取PPT文本内容到Markdown
- 提取图片到本地目录（按幻灯片索引命名）
- 自动修复图片链接为相对路径
- 图片OCR识别（支持中英文，使用easyocr）
- OCR噪声自动清理
- 图片与幻灯片精确对应

## 使用方法

### 基础用法

```powershell
python extract_pptx.py "presentation.pptx"
```

### 跳过OCR

```powershell
python extract_pptx.py "presentation.pptx" --no-ocr
```

### 指定输出目录

```powershell
python extract_pptx.py "presentation.pptx" "output_folder"
```

## 参数说明

| 参数 | 说明 | 必需 |
|------|------|------|
| `pptx_file` | PPTX文件路径 | 是 |
| `output_dir` | 输出目录（默认为PPTX同目录） | 否 |
| `--no-ocr` | 跳过图片OCR识别 | 否 |

## 输出格式

```markdown
<!-- Slide number: 1 -->1

![alt text](images/slide_1_5.png)
幻灯片文本内容...

---

**【幻灯片 1 - 图片识别内容】**

**图片OCR [slide_1_5.png]**

识别到的文字内容...
```

## 输出文件结构

```
{文件名}/                      # 与PPT同名的文件夹
├── {文件名}_含图片内容.md      # Markdown文件
└── images/                     # 提取的图片目录
    ├── slide_1_5.png          # 幻灯片1的第5个shape
    ├── slide_5_1.jpeg         # 幻灯片5的第1个shape
    └── ...
```

## 依赖包

脚本首次运行会自动安装以下依赖：

- `markitdown[pptx]` - PPTX文本提取
- `Pillow` - 图片处理
- `easyocr` - OCR识别
- `python-pptx` - PPTX解析
- `numpy` - 数组处理
- `opencv-python` - 图片读取（解决中文路径问题）

## 系统要求

- Python 3.8+
- 支持Windows/Linux/macOS

## 注意事项

1. 首次运行会自动安装依赖包
2. OCR使用easyocr，支持中英文识别
3. 图片按幻灯片索引命名，格式为 `slide_{幻灯片号}_{shape序号}.{扩展名}`
4. EMF格式图片不支持提取和OCR
5. 支持中文路径的图片文件
