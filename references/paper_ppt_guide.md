# 科研论文汇报 PPT 制作规范

## 一、整体原则

- **简洁清晰**：每页文字不超过 80 字，优先用图表、公式替代大段文字
- **图表优先**：所有图表必须严格来自原论文（截图或精确重绘），不得自行创造
- **结构完整**：必须覆盖论文的核心脉络，不遗漏关键贡献
- **页数控制**：含封面共 6–10 页
- **模板优先**：有参考模板时，所有视觉规格从模板中提取，本文默认值仅作兜底

---

## 二、输入材料说明

### 2.1 参考 PPT 模板（可选，强烈推荐）

用户提供的 .pptx 文件，用于提取视觉风格。

**从模板中提取的信息：**

| 信息类型 | 提取方式 | 用途 |
|---------|---------|------|
| 背景色 | 幻灯片背景填充色 | 每页背景 |
| 标题栏颜色 | 顶部色块的填充色 | 标题区背景 |
| 标题栏高度 | 色块高度 / 幻灯片高度 | 标题区布局 |
| 强调色 | 装饰线/重点色块的颜色 | 装饰线、关键词高亮 |
| 正文字体 | 文字框中的字体名称 | 所有正文 |
| 标题字体 | 标题文字框的字体名称 | 所有标题 |
| 字号 | 文字框字体大小 | 分级字号系统 |
| 内容布局 | shape 的相对位置和大小 | 确定双栏/全宽/卡片布局 |
| 装饰元素 | 线条、图标、分隔符 | 保持与模板一致的装饰风格 |
| 页脚 | 最底部的文字或线条 | 统一页脚格式 |

**注意：**
- 如果模板有多页，以**内容页**（非封面、非目录）为提取基准
- 颜色统一用 6 位 HEX（如 `1B2D5E`），不带 `#` 前缀，以匹配 python-pptx 的 `RGBColor.from_string()` 格式

### 2.2 讲解思路文档（可选，强烈推荐）

用户提供的演讲稿、汇报大纲或思路说明，用于确定每页结构和讲解重点。

**文档格式不限**：.txt / .md / .docx，或直接在对话中粘贴文字。

**从文档中提取的信息：**

| 信息类型 | 提取方式 | 用途 |
|---------|---------|------|
| 页面划分 | 文档是否按页/章节组织 | 确定 PPT 页数和章节顺序 |
| 每页讲解重点 | 每页要传达的核心 1–3 点 | 决定 bullet 内容 |
| 指定图表 | 文档提到的 Figure/Table | 优先裁剪对应图表 |
| 讲解顺序 | 章节顺序是否与标准结构一致 | 若有差异，以讲解文档为准 |
| 表述风格 | 口语/书面，中文/英文 | 统一 bullet 的表述风格 |

---

## 三、标准章节结构（无讲解文档时的默认结构）

| 序号 | 章节名称 | 内容要点 | 建议页数 |
|------|---------|---------|---------|
| 0 | 封面 | 论文标题、作者/机构、发表期刊/会议、汇报日期 | 1 页 |
| 1 | 问题背景与动机 | 研究背景、现有方法的不足、本文要解决的核心问题 | 1 页 |
| 2 | 重点概念解析 | 理解论文必须掌握的前置概念（只挑关键的）| 1 页 |
| 3 | 模型/方法架构 | 整体框架图（必须来自原文）、各模块功能说明 | 1–2 页 |
| 4 | 实验结果 | 主实验对比表/图（必须来自原文）、消融实验核心结论 | 1–2 页 |
| 5 | 创新点总结 | 3–5 条 bullet，对应论文 contribution | 1 页 |
| 6 | 局限性与未来方向 | 论文自述的局限 + 可能的后续工作 | 1 页 |

> 若论文无显式「重点概念」章节，可合并到背景页；若实验特别丰富，可拆为 2 页；总页数（含封面）不超过 10 页。

---

## 四、布局规范

### 4.1 有参考模板时

**完全以模板为准**，从模板中还原：
- 标题栏的高度、颜色、装饰线位置
- 内容区是双栏还是全宽还是卡片式
- 页脚格式

### 4.2 无参考模板时的默认布局

**PPT 尺寸**：
```
宽度: 13.33 英寸 (16:9标准)
高度: 7.50 英寸
```

**页面区域划分**：
```
┌──────────────────────────────────────┐
│           标题栏 (高度: 1.0英寸)       │  Y: 0 ~ 1.0
├──────────────────────────────────────┤
│                                      │
│                                      │
│         内容区域                      │  Y: 1.0 ~ 7.5
│      (高度: 6.5英寸)                  │
│                                      │
│                                      │
└──────────────────────────────────────┘
```

#### 模板1：左图右文（推荐，适合大多数内容页）

```
┌──────────────────────────────────────┐
│           标题栏 (1.0英寸)             │
├─────────────────┬────────────────────┤
│                 │                    │
│    图片区域      │     文字区域        │
│   宽度: 6.5英寸  │   宽度: 6.0英寸     │
│   高度: ≤5.5英寸 │   高度: ≤5.5英寸    │
│                 │                    │
│   X: 0.5        │   X: 7.0           │
│   Y: 1.3        │   Y: 1.3           │
└─────────────────┴────────────────────┘
```

**图片尺寸限制**：
- 最大宽度：6.5英寸
- 最大高度：5.5英寸
- 左边距：0.5英寸
- 上边距：1.3英寸（标题栏下方）

#### 模板2：右图左文

```
┌──────────────────────────────────────┐
│           标题栏 (1.0英寸)             │
├────────────────────┬─────────────────┤
│                    │                 │
│     文字区域        │    图片区域      │
│   宽度: 6.0英寸     │   宽度: 6.5英寸  │
│   高度: ≤5.5英寸    │   高度: ≤5.5英寸 │
│                    │                 │
│   X: 0.5           │   X: 7.0        │
│   Y: 1.3           │   Y: 1.3        │
└────────────────────┴─────────────────┘
```

#### 模板3：全宽图片（适用于大图/复杂架构图）

```
┌──────────────────────────────────────┐
│           标题栏 (1.0英寸)             │
├──────────────────────────────────────┤
│                                      │
│              图片区域                 │
│         宽度: 11.0英寸                │
│         高度: ≤5.5英寸                │
│                                      │
│         X: 1.0, Y: 1.3               │
└──────────────────────────────────────┘
```

#### 模板4：双栏文字（无图时使用）

```
┌──────────────────────────────────────┐
│           标题栏 (1.0英寸)             │
├──────────────────┬───────────────────┤
│                  │                   │
│    左栏内容       │    右栏内容        │
│   宽度: 6.0英寸   │   宽度: 6.0英寸    │
│   高度: ≤5.5英寸  │   高度: ≤5.5英寸   │
│                  │                   │
│   X: 0.5         │   X: 6.8          │
│   Y: 1.3         │   Y: 1.3          │
└──────────────────┴───────────────────┘
```

#### 封面布局

```
┌──────────────────────────────────────┐
│                                      │
│                                      │
│           论文标题                    │
│           （居中，28pt）              │
│                                      │
│           作者信息                    │
│           （居中，18pt）              │
│                                      │
│           期刊/会议 + 日期             │
│           （居中，14pt）              │
│                                      │
└──────────────────────────────────────┘
```

**布局选择原则**：
- 有图片时优先使用**左图右文**或**右图左文**
- 复杂架构图使用**全宽图片**
- 无图片时使用**双栏文字**

---

## 五、图表处理规范

### 5.1 图表来源要求（严格执行）

- **所有图表必须来自原论文**，包括架构图、实验对比表、消融实验图、可视化结果图
- **禁止**：自行绘制与原文不一致的架构图
- **禁止**：引用其他论文的图（除非原文明确引用且必须展示）
- **嵌入图片提取**：优先从PDF提取嵌入的原始图片，而非截图

### 5.2 图片提取流程

**第一步：提取嵌入图片**
```python
import fitz
import os

def extract_images_from_pdf(pdf_path, output_dir):
    """从PDF中提取所有嵌入图片"""
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
            
            image_filename = f"image_{page_num+1}_{img_index+1}.{image_ext}"
            image_path = os.path.join(output_dir, image_filename)
            
            with open(image_path, "wb") as f:
                f.write(image_bytes)
            
            image_list.append({
                "path": image_path,
                "page": page_num + 1,
                "size": len(image_bytes)
            })
    
    doc.close()
    return image_list
```

**第二步：按需裁剪区域**
```python
def crop_save(doc, page_idx, x0_pct, y0_pct, x1_pct, y1_pct, name, out_dir, dpi=280):
    """按百分比裁剪PDF页面区域"""
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
```

### 5.3 识别关键图片类型

1. **架构图/流程图**：用于方法介绍页（通常较大、包含多个模块、箭头连接）
2. **实验结果表/图**：用于实验结果页（包含数据对比、数字标注）
3. **可视化效果图**：用于定性分析页（展示效果对比，before/after）
4. **对比图**：用于方法对比页（左右或上下对比布局）
5. **消融实验图**：用于消融分析页（展示各模块贡献）

### 5.4 图片裁剪必须包含图注

**这是硬性要求，不可跳过：**
- 裁剪坐标必须延伸到图注文字结束处（下一节标题或空白行之前）
- 裁剪后必须用 `read_file` 工具查看裁剪结果
- 确认图注文字完整可见后才能使用
- 宁可多留 5% 空白，也不截断图注
- 子图的编号 (a)(b)(c)... 必须完整

### 5.5 图表选取优先级

1. **第一优先**：讲解文档中明确指定的图表
2. **第二优先**：论文的整体架构图（Overview Figure）
3. **第三优先**：主实验对比表（Main Results Table）
4. **第四优先**：消融实验中最能体现核心贡献的图/表
5. 其余可视化效果图（根据页数决定是否纳入）

### 5.6 图注标注

- 在图片下方添加来源注释：`(来源: 原文 Figure X / Table X)`
- 字号：12pt，颜色：灰色
- 若图表文字过小，在 bullet 中补充关键数字说明

### 5.7 图片命名规范

```
image_页码_序号.扩展名
例如：
- image_3_1.png（第3页第1张图）
- image_5_2.png（第5页第2张图）
- figure_1_overview.png（自定义命名）
```

---

## 六、文字规范

- **标题**：每页顶部，简明点题，≤ 15 字
- **Bullet 文字**：
  - 有讲解文档时：优先使用讲解文档中的表述
  - 无讲解文档时：从论文中提炼，动词开头，避免名词堆砌
  - 每条 ≤ 30 字，最多 5 条/页
- **技术术语**：首次出现时括号注明英文原文，例如：掩码自编码器（Masked AutoEncoder, MAE）
- **数字**：关键数字加粗，例如 **+2.3% mAP**

---

## 七、默认配色与风格（有参考模板时从模板提取，不用以下默认值）

### 配色方案（学术风格）

- **主色/标题栏**：深蓝 `#1E3A5F`（专业、稳重）
- **辅色/强调**：浅蓝 `#4A90D9`（用于重点提示）
- **背景**：白色 `#FFFFFF`
- **正文**：深灰 `#333333`
- **次要文字**：中灰 `#666666`
- **强调色**：橙色 `#E07040`（用于关键数字/创新点高亮）

### 字体规范

| 元素 | 字号 | 字重 |
|------|------|------|
| 标题 | 28pt | Bold |
| 小标题 | 18pt | Bold |
| 正文 | 16pt | Regular |
| 图注 | 12pt | Regular |
| 页码 | 10pt | Regular |

### 禁止事项

- ❌ 渐变背景
- ❌ 过多装饰性元素
- ❌ 超过 3 种主色
- ❌ 超过 5 种字体大小层级

---

## 八、pptx 生成代码示例

### 基础设置

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# 创建PPT
prs = Presentation()
prs.slide_width = Inches(13.333)  # 16:9 标准
prs.slide_height = Inches(7.5)

# 默认配色（来自第七节）
DEFAULT_COLORS = {
    "title_bar": RGBColor(0x1E, 0x3A, 0x5F),  # 深蓝 #1E3A5F
    "accent": RGBColor(0x4A, 0x90, 0xD9),      # 浅蓝 #4A90D9
    "body": RGBColor(0x33, 0x33, 0x33),       # 深灰 #333333
    "subtext": RGBColor(0x66, 0x66, 0x66),    # 中灰 #666666
    "highlight": RGBColor(0xE0, 0x70, 0x40),  # 橙色 #E07040
}
```

### 添加标题栏

```python
def add_title_bar(slide, title_text, colors=DEFAULT_COLORS):
    """添加标题栏（深蓝色块 + 白色标题文字）"""
    # 标题栏背景
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.0)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = colors["title_bar"]
    header.line.fill.background()  # 无边框

    # 标题文字
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.333), Inches(0.6))
    tf = title_box.text_frame
    tf.paragraphs[0].text = title_text
    tf.paragraphs[0].font.size = Pt(28)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
```

### 添加左图右文布局

```python
def add_left_image_right_text(slide, image_path, text_items, title_text="", colors=DEFAULT_COLORS):
    """左图右文布局
    Args:
        image_path: 图片路径
        text_items: 文字内容列表 ["标题", "• 要点1", "• 要点2", ...]
        title_text: 页面标题（可选）
    """
    if title_text:
        add_title_bar(slide, title_text, colors)

    # 添加图片（左侧，宽度6.5英寸）
    if image_path:
        slide.shapes.add_picture(
            image_path,
            Inches(0.5),   # X坐标
            Inches(1.3),   # Y坐标（标题栏下方）
            width=Inches(6.5)
        )

    # 添加文字（右侧）
    text_box = slide.shapes.add_textbox(
        Inches(7.0),   # X坐标（图片右侧）
        Inches(1.3),    # Y坐标
        Inches(6.0),    # 宽度
        Inches(5.5)     # 高度
    )
    tf = text_box.text_frame
    tf.word_wrap = True

    for i, line in enumerate(text_items):
        if i == 0:
            p = tf.paragraphs[0]
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = colors["title_bar"]
        else:
            p = tf.add_paragraph()
            p.text = line
            p.font.size = Pt(16)
            p.font.color.rgb = colors["body"]
        p.space_after = Pt(6)
```

### 添加图注

```python
def add_figure_caption(slide, caption_text, x, y, colors=DEFAULT_COLORS):
    """添加图注
    Args:
        caption_text: 图注内容，如 "(来源: 原文 Figure 1)"
        x, y: 位置坐标（英寸）
    """
    caption = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(6.5), Inches(0.3))
    tf = caption.text_frame
    tf.paragraphs[0].text = caption_text
    tf.paragraphs[0].font.size = Pt(12)
    tf.paragraphs[0].font.color.rgb = colors["subtext"]
```

---

## 九、常规注意事项

- 使用 `python-pptx` 库生成 .pptx 文件
- 图表用 `add_picture` 插入，保持原始宽高比
- 文字框用 `add_textbox`，所有字号从 `DESIGN` 字典读取
- 幻灯片尺寸：宽 33.87cm × 高 19.05cm（16:9 标准）
- **有模板时**：`DESIGN` 字典完全从模板解析结果填充，不使用本文默认值
- **无模板时**：使用本文第七节的默认配色与风格
