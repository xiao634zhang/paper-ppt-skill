---
name: paper-ppt
description: >
  This skill should be used whenever the user wants to create a PowerPoint presentation (科研汇报PPT)
  for a single academic research paper. Triggers include: "帮我做这篇论文的汇报PPT"、"把这篇paper做成PPT"、
  "论文PPT"、"paper presentation"、"科研汇报"、"文献汇报". The skill produces a concise, figure-first
  6–10 slide .pptx file that covers background, key concepts, model architecture, experimental results,
  contributions, and future directions—strictly using figures and tables from the original paper.
  Supports user-provided reference template (.pptx) and speaker notes / outline document for style and structure alignment.
---

# 科研论文汇报 PPT Skill

## 目的

将单篇科研论文转化为简洁清晰、图表优先的汇报 PPT（.pptx 格式），适用于组会汇报、课程展示、学术交流等场景。

**支持用户提供两个可选的参考文件，以提升生成质量：**
1. **参考 PPT 模板**（.pptx 文件）：提供视觉风格参考，包括配色、字体、布局、装饰元素
2. **讲解思路文档**（.txt / .md / .docx 文件，或直接粘贴文字）：提供汇报的结构、讲解逻辑和每页要讲的重点

---

## 工作流程

### Step 0：收集输入材料

在开始之前，**向用户询问以下材料**（均为可选，但强烈推荐）：

```
我需要以下材料来制作 PPT：

① 论文文件（必填）：PDF 文件路径
② 参考 PPT 模板（可选，强烈推荐）：一份你觉得排版好看的 .pptx 文件
   → 我会从中提取配色、字体、布局风格，照着做
③ 讲解思路 / 演讲稿（可选，强烈推荐）：说明每页要讲什么、讲解顺序和重点
   → 可以是文本文件，也可以直接告诉我大纲
```

如果用户已经提供了部分材料（比如直接给了 PDF 路径），就不要重复问已有的部分，只问缺的。

---

### Step 1：解析参考 PPT 模板（如有）

若用户提供了参考 .pptx 文件，执行以下解析流程：

#### 1.1 用 python-pptx 读取模板结构

```python
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import json

prs = Presentation(template_path)

# 提取幻灯片尺寸
slide_w = prs.slide_width.cm
slide_h = prs.slide_height.cm

# 遍历所有幻灯片，提取颜色、字体、布局信息
style_info = []
for i, slide in enumerate(prs.slides):
    page = {"page": i+1, "shapes": []}
    for shape in slide.shapes:
        item = {
            "type": shape.shape_type,
            "name": shape.name,
            "left_pct": round(shape.left / prs.slide_width, 3),
            "top_pct": round(shape.top / prs.slide_height, 3),
            "width_pct": round(shape.width / prs.slide_width, 3),
            "height_pct": round(shape.height / prs.slide_height, 3),
        }
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size:
                        item["font_size_pt"] = round(Pt(1).pt and run.font.size.pt, 1)
                    if run.font.color and run.font.color.type:
                        try:
                            item["font_color"] = str(run.font.color.rgb)
                        except:
                            pass
                    if run.font.bold:
                        item["bold"] = True
                    if run.font.name:
                        item["font_name"] = run.font.name
        # 提取填充色
        try:
            fill = shape.fill
            if fill.type is not None and fill.fore_color:
                item["fill_color"] = str(fill.fore_color.rgb)
        except:
            pass
        page["shapes"].append(item)
    style_info.append(page)

print(json.dumps(style_info[:3], ensure_ascii=False, indent=2))  # 先看前3页
```

#### 1.2 截图目视确认

```python
# 将参考PPT的每一页渲染为图片，用 read_file 工具逐页查看
# 方法：用 win32com (Windows) 或 LibreOffice (Linux) 转 PDF，再用 pymupdf 渲染
# 若环境不支持，可要求用户提供截图或跳过此步
```

**⚠️ 目视确认优先**：先看截图，理解模板的整体视觉风格，再看代码解析数据。
- 记录：背景色、主标题颜色、装饰条/色块的颜色和位置、内容区布局（单栏/双栏/卡片）
- 记录：字体风格（衬线/无衬线）、是否有页眉页脚、是否有图标或装饰元素

#### 1.3 提炼设计规格

将解析结果整理成如下格式，供后续生成脚本使用：

```python
DESIGN = {
    # 颜色（从模板中提取，找出主色、背景色、强调色）
    "bg_color": "FFFFFF",        # 幻灯片背景色
    "title_bar_color": "1B2D5E", # 标题栏/顶部色块
    "accent_color": "E07040",    # 强调色（装饰线/重点文字）
    "body_text_color": "333333", # 正文文字颜色
    "subtext_color": "666666",   # 次要文字/图注颜色

    # 字体（从模板中提取）
    "font_title": "Arial",       # 标题字体
    "font_body": "Arial",        # 正文字体
    "font_size_title": 28,       # 标题字号
    "font_size_body": 18,        # 正文字号
    "font_size_caption": 12,     # 图注字号

    # 布局（从模板中提取）
    "title_bar_height_pct": 0.15, # 顶部标题栏高度占比
    "has_decoration_line": True,  # 是否有竖线/横线装饰
    "decoration_line_color": "E07040",
    "content_top_pct": 0.18,     # 内容区起始位置
    "has_footer": True,          # 是否有底部页脚

    # 幻灯片尺寸
    "slide_w_cm": 33.87,
    "slide_h_cm": 19.05,
}
```

---

### Step 2：解析讲解思路文档（如有）

若用户提供了演讲稿或大纲文档，执行以下流程：

#### 2.1 读取文档内容

```python
# .txt / .md 文件：直接 read_file 读取
# .docx 文件：用 python-docx 提取文字
from docx import Document
doc = Document(outline_path)
text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
print(text)
```

#### 2.2 提取结构信息

从文档中识别：
- **页面划分**：文档是否已经按页/章节组织？每页对应哪个主题？
- **每页讲解重点**：每页要传达的 1-3 个核心信息
- **讲解顺序**：是否与标准结构（背景→方法→实验→总结）一致？如有差异，以讲解文档为准
- **指定图表**：文档是否提到要用哪个 Figure/Table？

整理为结构化大纲：

```
页面结构（来自讲解文档）：
1. 封面 → [论文标题 + 作者]
2. 背景 → 重点：CT辐射剂量问题 + 现有方法缺陷
3. 方法概览 → 重点：可微渲染框架 + Beer-Lambert
...
```

#### 2.3 与论文内容对齐

将讲解文档的结构与从 PDF 提取的内容对照，确认：
- 讲解文档提到的图表是否都能在 PDF 中找到
- 讲解文档的描述是否准确（避免讲解稿本身有误）

---

### Step 3：读取论文内容

**⚠️ 严格要求：必须逐页读取 PDF 截图，用眼睛确认论文标题和内容，再开始提取。**

- 先读取第 1 页截图确认：论文标题、作者、期刊/会议，**避免同名论文混淆**
- 再逐页读取截图，了解全文结构和图表位置

---

### Step 3.1：从 PDF 提取嵌入图片（重要！）

**如果论文中有嵌入的 Figure 图片，必须提取出来！**

```python
import fitz  # PyMuPDF
import os

def extract_images_from_pdf(pdf_path, output_dir):
    """从PDF中提取所有嵌入图片
    
    Returns:
        list: [{"path": 图片路径, "page": 页码, "xref": xref编号, "size": 文件大小}]
    """
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    
    image_list = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        images = page.get_images(full=True)
        
        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            
            # 图片命名规范：image_页码_序号.扩展名
            image_filename = f"image_{page_num+1}_{img_index+1}.{image_ext}"
            image_path = os.path.join(output_dir, image_filename)
            
            with open(image_path, "wb") as f:
                f.write(image_bytes)
            
            image_list.append({
                "path": image_path,
                "page": page_num + 1,
                "xref": xref,
                "size": len(image_bytes)
            })
    
    doc.close()
    return image_list

# 使用示例
pdf_path = r"path/to/paper.pdf"
out_dir = r"path/to/figures"
images = extract_images_from_pdf(pdf_path, out_dir)
for img in images:
    print(f"提取: {img['path']} (第{img['page']}页, {img['size']/1024:.1f}KB)")
```

**识别关键图片类型**：
1. **架构图/流程图**：用于方法介绍页（通常较大、包含多个模块）
2. **实验结果表/图**：用于实验结果页（包含数据对比）
3. **可视化效果图**：用于定性分析页（展示效果对比）
4. **对比图**：用于方法对比页（左右或上下对比布局）

---

### Step 3.2：逐页渲染 PDF 截图

```python
import fitz
import os

pdf_path = r"path/to/paper.pdf"
out_dir = r"path/to/figures"
os.makedirs(out_dir, exist_ok=True)

doc = fitz.open(pdf_path)
# 逐页渲染为图片，dpi=200 足够阅读
for i, page in enumerate(doc):
    mat = fitz.Matrix(200/72, 200/72)
    pix = page.get_pixmap(matrix=mat)
    pix.save(f"{out_dir}/page_{i+1}.png")

doc.close()
print(f"共 {len(doc)} 页，已保存到 {out_dir}")
```

用 `read_file` 工具逐页读取 PNG，记录：
- 论文标题（第1页）
- 每个 Figure/Table 在哪一页、大致位置（上/中/下）
- 关键文字内容（Abstract、Contributions、Conclusion）

---

### Step 4：提取核心要素

结合论文内容和讲解思路文档（如有），提取以下内容：

1. **研究问题**：用 1 句话描述论文要解决什么问题
2. **关键前置概念**：理解论文必须掌握的 2–4 个概念（来自论文本身）
3. **核心架构**：整体方法框图（记录是哪个 Figure）
4. **主要实验结果**：最重要的 1–2 张对比表/图（记录 Table/Figure 编号）
5. **创新点**：直接引用论文 contribution 列表
6. **局限/未来**：论文 conclusion 中的 limitation 和 future work

> 若有讲解思路文档，以文档中的重点为准，从论文中找对应内容；若无，按标准流程提取。

---

### Step 5：裁剪图表

从 PDF 裁剪所需图表，**每张图必须包含完整图注**：

```python
def crop_save(doc, page_idx, x0_pct, y0_pct, x1_pct, y1_pct, name, out_dir, dpi=280):
    """按百分比裁剪PDF页面区域并保存为图片
    
    Args:
        doc: fitz.Document 对象
        page_idx: 页码（从0开始）
        x0_pct, y0_pct: 左上角百分比坐标（0.0 ~ 1.0）
        x1_pct, y1_pct: 右下角百分比坐标（0.0 ~ 1.0）
        name: 输出文件名（不含扩展名）
        out_dir: 输出目录
        dpi: 输出分辨率，默认280
    """
    page = doc[page_idx]
    r = page.rect
    x0 = r.x0 + r.width  * x0_pct
    y0 = r.y0 + r.height * y0_pct
    x1 = r.x0 + r.width  * x1_pct
    y1 = r.y0 + r.height * y1_pct
    clip = fitz.Rect(x0, y0, x1, y1)
    mat = fitz.Matrix(dpi/72, dpi/72)
    pix = page.get_pixmap(matrix=mat, clip=clip)
    path = f"{out_dir}/{name}.png"
    pix.save(path)
    return path

# 使用示例：裁剪第3页的图片区域（根据观察到的位置调整百分比）
# doc = fitz.open("paper.pdf")
# crop_save(doc, page_idx=2, x0_pct=0.05, y0_pct=0.1, x1_pct=0.95, y1_pct=0.8, 
#           name="figure_1_overview", out_dir="figures", dpi=280)
```

**快速裁剪法（适合规则分布的图表）**：

```python
def crop_save(doc, page_idx, x0_pct, y0_pct, x1_pct, y1_pct, name, out_dir, dpi=280):
    """按百分比裁剪PDF页面区域并保存为图片
    
    Args:
        doc: fitz.Document 对象
        page_idx: 页码（从0开始）
        x0_pct, y0_pct: 左上角百分比坐标（0.0 ~ 1.0）
        x1_pct, y1_pct: 右下角百分比坐标（0.0 ~ 1.0）
        name: 输出文件名（不含扩展名）
        out_dir: 输出目录
        dpi: 输出分辨率，默认280
    """
    page = doc[page_idx]
    r = page.rect
    x0 = r.x0 + r.width  * x0_pct
    y0 = r.y0 + r.height * y0_pct
    x1 = r.x0 + r.width  * x1_pct
    y1 = r.y0 + r.height * y1_pct
    clip = fitz.Rect(x0, y0, x1, y1)
    mat = fitz.Matrix(dpi/72, dpi/72)
    pix = page.get_pixmap(matrix=mat, clip=clip)
    path = f"{out_dir}/{name}.png"
    pix.save(path)
    return path

# 使用示例：裁剪第3页的图片区域（根据观察到的位置调整百分比）
# doc = fitz.open("paper.pdf")
# crop_save(doc, page_idx=2, x0_pct=0.05, y0_pct=0.1, x1_pct=0.95, y1_pct=0.8, 
#           name="figure_1_overview", out_dir="figures", dpi=280)
```

**快速裁剪法（适合规则分布的图表）**：
```python
def crop_all_figures_from_page(doc, page_idx, left_pct, top_pct, right_pct, bottom_pct, 
                                out_dir, prefix="fig", dpi=280):
    """从一页中按固定区域批量裁剪多张图
    
    适用于同一页有多个并排/上下分布的子图
    """
    page = doc[page_idx]
    r = page.rect
    
    # 计算裁剪区域
    x0 = r.x0 + r.width * left_pct
    y0 = r.y0 + r.height * top_pct
    x1 = r.x0 + r.width * right_pct
    y1 = r.y0 + r.height * bottom_pct
    
    clip = fitz.Rect(x0, y0, x1, y1)
    mat = fitz.Matrix(dpi/72, dpi/72)
    pix = page.get_pixmap(matrix=mat, clip=clip)
    
    path = f"{out_dir}/{prefix}_page{page_idx+1}.png"
    pix.save(path)
    return path
```

**裁剪后必须用 `read_file` 工具逐张查看**，确认：
- ✅ 图注文字完整（不能被截断）
- ✅ 图片内容清晰完整
- ✅ 子图编号（a)(b)(c)... 完整可见
- 若图注截断：扩大 `y1_pct`，宁可多留 5% 空白也不截断
- 若子图截断：调整 left/right/top/bottom 百分比

**图片命名规范**：
```
image_页码_序号.扩展名
例如：image_3_1.png（第3页第1张图）、image_5_2.png（第5页第2张图）
```

---

### Step 6：规划 PPT 结构

**优先级**：
1. 若有讲解思路文档 → 以文档的页面划分和讲解重点为准
2. 若无 → 按以下默认结构（可根据论文内容调整）

```
第 1 页  封面        论文标题、作者、期刊/会议、汇报日期
第 2 页  问题背景    研究背景 + 现有方法不足 + 本文目标
第 3 页  重点概念    2–4 个关键概念解析（如无独特概念可合并到背景）
第 4 页  模型架构    Overview Figure + 各模块说明（大图时可拆 2 页）
第 5 页  实验结果    主实验 Table/Figure + 关键数字标注
第 6 页  消融实验    消融实验图表（如内容少可与第5页合并）
第 7 页  创新点总结  3–5 条 bullet，直接对应论文 contribution
第 8 页  局限与展望  论文自述局限 + 未来方向
```

> 总页数（含封面）≤ 10 页

---

### Step 7：生成 .pptx 文件

使用 `python-pptx` 生成幻灯片。

#### 样式来源优先级（严格遵守）

```
有参考模板 → 从 DESIGN 字典读取所有样式参数（颜色/字体/布局）
无参考模板 → 使用 paper_ppt_guide.md 中的默认规范
```

**绝不允许**在有参考模板时仍使用默认配色。

#### 布局选择

- **主布局（大多数内容页）**：根据模板风格决定（双栏 / 全宽 / 卡片）
- **有讲解文档时**：每页的 bullet 内容从讲解文档对应章节中提取，不自行发挥
- **封面页**：从模板封面提取布局，没有模板则居中排版

#### 图表插入规则（严格执行）

- 所有图表**必须来自原论文**，使用 `add_picture` 插入裁剪好的图片
- 每张图下方添加图注（来自论文原文图注的精简版，≤ 20 字）
- 若无法获取图片，用文字框标注「[原文 Figure X]」占位

#### 文字规则

- 标题：`DESIGN["font_size_title"]` pt，粗体
- 正文 bullet：`DESIGN["font_size_body"]` pt，每页 ≤ 5 条，每条 ≤ 30 字
- 图注：`DESIGN["font_size_caption"]` pt，灰色
- **若有讲解文档**：bullet 内容优先取讲解文档中的表述，而非自行总结

---

### Step 8：输出与交付

- 将 .pptx 保存到用户指定目录，或与论文同目录
- 文件命名：`[论文短标题]_汇报PPT.pptx`
- 输出后说明：
  - 每页内容来源（哪页用了哪个 Figure/Table）
  - 是否有图片占位符需手动替换
  - 若有讲解文档，说明哪些地方采纳了讲解文档的表述

---

## 注意事项

- **图注截断禁令**：每次裁图后必须目视确认图注完整，这是硬性要求
- **同名论文陷阱**：读 PDF 第一步必须确认标题，避免混淆
- **模板优先**：有参考模板时，所有颜色/字体/装饰从模板中提取，不用默认值覆盖
- **讲解文档优先**：有讲解文档时，每页的结构和重点以文档为准，论文内容为素材来源
- **不得自行创造图表**：若无法从论文获取图片，使用文字占位符
- **不得修改论文结论**：创新点和实验结论需忠实原文
- **语言一致性**：若论文是英文，PPT 可中英混排；若用户明确要求纯中文，翻译关键术语
- **详细排版默认参数**：参见 `references/paper_ppt_guide.md`
- **嵌入图片优先**：优先提取PDF中嵌入的原始图片，而非截图

---

## 检查清单

制作完成后，检查以下项目：

- [ ] 是否从PDF提取了嵌入图片？
- [ ] 逐页渲染截图确认论文标题了吗？
- [ ] 有图的页面是否优先放了图？
- [ ] 图片大小是否超出页面限制？
- [ ] 文字内容是否超出内容区域？
