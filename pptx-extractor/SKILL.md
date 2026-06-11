# PPTX Extractor Skill

将PPTX文件解析为Markdown格式，支持图片OCR识别和图片提取。

## 功能特点

- ✅ 提取PPT文本内容到Markdown
- ✅ **提取图片到本地目录（支持PNG/TIFF等格式）**
- ✅ **修复图片链接为有效相对路径**
- ✅ 图片OCR识别（支持中英文）
- ✅ 时间轴内容智能整理
- ✅ OCR噪声自动清理
- ✅ 图片与幻灯片对应显示

## 快速开始

### 方法一：双击安装（推荐）

1. 运行 `install.bat`
2. 按照提示完成安装
3. 使用 `使用PPTX提取器.bat` 处理PPTX文件

### 方法二：命令行安装

```powershell
pip install markitdown[pptx] Pillow pytesseract python-pptx
```

### 方法三：手动安装Tesseract

下载地址：https://github.com/UB-Mannheim/tesseract/wiki

## 使用方法

### 基础用法

```powershell
python extract_pptx.py "presentation.pptx"
```

### 指定路径

```powershell
python extract_pptx.py "presentation.pptx" ^
    --tesseract "D:\Program\Tesseract-OCR\tesseract.exe"
```

### 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--python` | Python可执行文件路径 | 系统Python |
| `--tesseract` | Tesseract路径 | PATH中的tesseract |
| `--no-ocr` | 跳过图片OCR | 否 |
| `--no-semantic` | 跳过语义优化 | 否 |

## 输出格式

```markdown
<!-- Slide number: 2 -->2
【时空坐标·历史脉络】

![CM14.TIF](images/image1.png)

---

**【幻灯片 2 - 图片识别内容】**

**图片OCR [image.png]**

1922年 → 1931年 → 1933年 → 1936年 → 1938年 → 1939年 → 1942年 → 1945年
```

## 输出文件结构

```
{文件名}/                    # 与PPT同名的文件夹
├── {文件名}_含图片内容.md    # Markdown文件（带图片链接）
└── images/                   # 提取的图片目录
    ├── image1.png
    ├── image2.png
    └── ...
```

**多个PPT处理示例：**

```
输出目录/
├── 第17课/
│   ├── 第17课_含图片内容.md
│   └── images/
├── 第18课/
│   ├── 第18课_含图片内容.md
│   └── images/
└── 第19课/
    ├── 第19课_含图片内容.md
    └── images/
```

每个PPT都有独立的文件夹，互不干扰。

## 系统要求

- Windows 10/11
- Python 3.8+
- Tesseract OCR（可选，用于图片识别）

## 注意事项

1. 首次运行会自动安装依赖包
2. 建议将Tesseract添加到系统PATH
3. 输出文件保存在PPTX同一目录下
4. EMF格式图片不支持OCR识别