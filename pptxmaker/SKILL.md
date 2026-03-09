---
name: ppt-maker
description: >
  Create, edit, and read PowerPoint presentations using python-pptx.
  Use this skill when the user asks to create slides, make a presentation,
  generate a PPT/PPTX, build a deck, or work with PowerPoint files.
  Handles template-based creation, chart insertion, and slide editing.
  Also triggered by: 幻灯片, 演示文稿, 做PPT, 生成PPT, 编辑PPT, 读取PPT.
---

# PPT Maker Skill

你是 PPT 创建与编辑专家。使用 `python-pptx` 库通过 Python 脚本生成和编辑 PowerPoint 文件。

## 核心原则

1. **始终生成原生可编辑的 PPTX**，而非图片式幻灯片
2. **使用辅助库** `scripts/ppt_helpers.py` 简化操作
3. **严格遵守品牌规范**：必须使用内置的 `bank_brand` 品牌颜色和特定字体
4. **母版保护机制**：对于使用企业模板创建的 PPT，**封面页、结束页、分隔页的版式原则上不允许修改**。只能修改这些页面上的文本内容，绝不更改布局。
5. **字体要求**：中文严格使用 **"圆体-简"**，数字和英文严格使用 **"effra"** 字体。

## 默认配置

- 默认配色方案：`bank_brand`（基础 RGB：红 255/85/5，黄 244/185/0，紫 53/4/64）
- 默认语言：`zh-CN`
- 默认中文字体：`Microsoft YaHei`
- 默认英文字体：`Calibri`
- 输出文件名格式：`{主题}_{YYYYMMDD}.pptx`

## 环境准备

在首次使用前，请确保已安装 python-pptx 依赖：

```bash
pip install -r scripts/requirements.txt
```

验证安装：
```bash
python3 -c "import pptx; print(f'✅ python-pptx {pptx.__version__} 安装验证完成!')"
```

---

## 使用场景

### 场景 1：创建 PPT

根据用户的自然语言描述生成 PowerPoint 演示文稿。

**步骤：**
1. 分析请求以确定：主题、幻灯片数量、风格/配色偏好
2. 从配色方案中选择合适配色（默认：`bank_brand`）
3. 设计幻灯片结构（为每页选择相应的布局类型）
4. 编写使用 python-pptx 的 Python 脚本来生成 PPTX。使用 `scripts/ppt_helpers.py` 中的辅助函数
5. 执行该脚本
6. 将文件保存到当前工作目录，格式为：`{主题}_{YYYYMMDD}.pptx`
7. 报告输出文件路径及创建内容的摘要

**规则：**
- 使用 `sys.path.insert(0, '<skill_dir>/scripts')` 导入辅助库（其中 `<skill_dir>` 是本 SKILL.md 所在目录的绝对路径）
- 每页幻灯片不应超过 6 条项目符号要点
- 标题字体大小：28-36pt，正文字体大小：18-24pt
- 在整个 PPT 中始终应用统一的配色方案

---

### 场景 2：编辑 PPT

编辑现有的 PowerPoint 文件（添加/删除/修改幻灯片页面）。

**步骤：**
1. 解析用户需求，提取文件路径和编辑指令
2. 在开始编辑前，先运行以下命令读取并分析现有文件结构：
   ```bash
   python3 scripts/read_ppt.py "<文件路径>"
   ```
3. 检查控制台解析输出以了解原 PPT 的幻灯片和结构
4. 备份原始文件为 `{原文件名}_backup.pptx`
5. 根据要求的编辑变更，使用 python-pptx 编写 Python 脚本
6. 使用 `sys.path.insert(0, '<skill_dir>/scripts')` 导入辅助库
7. 保存修改后的文件
8. 在结果反馈中汇总已变更的内容列表

**可支持的编辑操作：**
- 添加新幻灯片（需指定插入的位置和所需内容）
- 指定页码或索引删除特定幻灯片
- 替换或修改现有幻灯片的标题及正文内容
- 更新字体、颜色和其他排版相关样式
- 更改幻灯片的先后排列顺序
- 删除或加入图像、Excel 图表、以及表格区域

---

### 场景 3：从模板创建

根据 `resources/templates/` 目录下的现有模板文件创建演示文稿。

**步骤：**
1. 解析用户需求，提取模板名称和主题描述
2. 在 `resources/templates/{模板名称}.pptx` 处定位该模板文件
3. 如果未找到对应模板，列出支持的可用模板并请用户选择
4. 使用 `scripts/read_ppt.py` 检查并分析该模板的母版布局
5. 编写 Python 脚本打开该模板，并在保持原风格的基础上生成内容
6. **硬性规定 — 母版保护机制**：绝不能修改封面页、结束页或分隔页的布局、位置或格式规范。只允许基于占位符替换文本。
7. **硬性规定 — 字体规范要求**：确保所有中文文本使用"圆体-简"字体，所有英文字母与数字使用"effra"字体。
8. 以描述性的名称和日期保存文件：`{主题}_{YYYYMMDD}.pptx`

#### 关键技术规范（必须遵守）

##### 规范 1：使用 duplicate_slide 复制模板 slide，禁止使用 add_slide(layout)

模板 slide 中的文本框、图表、图片等绑定在 slide 本身，而非 slide_layout 上。
`prs.slides.add_slide(layout)` 只会继承 layout 中的占位符，不会复制那些手动添加的 shapes。

**正确做法**：使用 `ppt_helpers.duplicate_slide(prs, src_slide_index)` 深拷贝源 slide，
然后遍历已复制 slide 的 shapes 进行文本替换。

```python
from ppt_helpers import duplicate_slide, set_run_font

# 复制模板的第3张 slide（索引2）
new_slide = duplicate_slide(prs, 2)
# 遍历 shapes 找到目标文本框并替换
for shape in new_slide.shapes:
    if shape.has_text_frame and shape.name == '文本框 7':
        # 替换文本...
```

##### 规范 2：智能双字体设置 — 使用 set_run_font()

Python-pptx 中 `run.font.name` 是拉丁字体属性，PowerPoint/WPS 对中文文本优先使用东亚字体(`a:ea`)。

**正确做法**：使用 `ppt_helpers.set_run_font(run, font_cn, font_en)` 自动判断：
- 含中文字符 → `font.name = font_cn`（如"圆体-简"），同时 `a:ea = font_cn`
- 纯英文/数字 → `font.name = font_en`（如"effra"），同时 `a:ea = font_cn`（回退用）

```python
from ppt_helpers import set_run_font

for para in text_frame.paragraphs:
    for run in para.runs:
        set_run_font(run, font_cn='圆体-简', font_en='effra')
```

##### 规范 3：替换文本时保留原始格式

替换已复制 slide 中的文本时，应尽量保留原始 run 的格式属性（字号、颜色、粗体等），
只替换 `run.text` 的值，然后用 `set_run_font()` 修正字体。

```python
for para in shape.text_frame.paragraphs:
    for run in para.runs:
        run.text = '新文本'
        set_run_font(run, '圆体-简', 'effra')
```

##### 规范 4：正文页标题必须对齐设计规范

模板中不同正文页的标题 shape 位置各不相同，复制 slide 后必须使用 `fix_content_title()` 将标题统一归位。

**标题规范**: 正文的标题需要放在左上角，在长方形图标之后，在横线之上，并把正文标题的字体设置成24号。

```python
from ppt_helpers import fix_content_title

title_shape = find_shape(slide, '文本框 3')
fix_content_title(title_shape, '新标题文本')
```

---

### 场景 4：读取 PPT

提取、读取并解析给定的 PowerPoint 文件的整体结构和包含的文本内容。

运行以下命令执行读取操作：
```bash
python3 scripts/read_ppt.py "<PPT文件路径>"
```

提取 JSON 输出，并以清晰、易读且符合人类阅读习惯的 Markdown 格式展示结果：
- 文档总体信息：幻灯片总数、页面宽高及尺寸
- 针对每一页单独总结：序号/索引号、对应母版布局名、主体文本内容、相关配图数目、内置图表、和表格数量等信息
- 需要重点强调任何可以识别到的模板特性或显著规律

---

## python-pptx API 速查

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData

# 创建/打开
prs = Presentation()                        # 新建
prs = Presentation('template.pptx')         # 从模板

# 设置尺寸 (16:9)
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# 添加幻灯片
slide_layout = prs.slide_layouts[6]         # 空白布局
slide = prs.slides.add_slide(slide_layout)

# 文本框
txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
tf = txBox.text_frame
tf.text = "标题文字"
tf.paragraphs[0].font.size = Pt(28)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)

# 图片
slide.shapes.add_picture('image.png', Inches(1), Inches(2), Inches(6), Inches(4))

# 表格
table = slide.shapes.add_table(3, 4, Inches(1), Inches(2), Inches(8), Inches(3)).table

# 图表
chart_data = CategoryChartData()
chart_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
chart_data.add_series('收入', (100, 200, 300, 400))
slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                       Inches(1), Inches(2), Inches(10), Inches(5), chart_data)

# 保存
prs.save('output.pptx')
```

## 预设布局系统

8 种常用布局（详细坐标参数见 `resources/layouts.md`）：

| 布局 | 方法名 | 适用场景 |
|------|--------|---------|
| 标题页 | `add_title_slide()` | 封面 |
| 内容页 | `add_content_slide()` | 标准要点页 |
| 两栏页 | `add_two_column_slide()` | 对比/并列 |
| 图文页 | `add_image_slide()` | 图片展示 |
| 图表页 | `add_chart_slide()` | 数据可视化 |
| 章节页 | `add_section_slide()` | 章节过渡 |
| 总结页 | `add_summary_slide()` | 关键要点 |

## 专业配色方案

6 套配色（详细色值见 `resources/color-palettes.md`）：

| 主题 | 名称 | 适用 |
|------|------|------|
| 商务蓝 | `business_blue` | 企业汇报、年报 |
| 科技感 | `tech_dark` | 技术分享、产品发布 |
| 清新绿 | `fresh_green` | 教育、健康 |
| 暖橙 | `warm_orange` | 营销、创意 |
| 优雅紫 | `elegant_purple` | 设计、文化 |
| 简约灰 | `minimal_gray` | 通用极简 |

## 设计规则

- 每页要点 **不超过 6 条**
- 标题字号 **28-36pt**，正文 **18-24pt**
- 全局统一 **默认使用 `bank_brand` 品牌配色方案**
- 字体严格遵循：中文使用 **"圆体-简"**，英文/数字使用 **"effra"**
- 封面页、结束页、章节标题页的样式必须继承模板母版，不可破坏原有版式结构
- 图表优先使用 **原生 PowerPoint 图表对象**，且图表颜色需应用品牌的基础色和辅助色
- 输出文件名格式：`{主题}_{YYYYMMDD}.pptx`

## 辅助库使用指南

### 使用 PPTHelper 类

```python
import sys, os
skill_dir = os.path.dirname(os.path.abspath(__file__))  # 或者使用已知的 skill 目录路径
sys.path.insert(0, os.path.join(skill_dir, 'scripts'))
from ppt_helpers import PPTHelper

helper = PPTHelper(theme="business_blue")
helper.add_title_slide("演示标题", "副标题")
helper.add_content_slide("要点", ["点1", "点2", "点3"])
helper.add_chart_slide("数据", "bar", ["Q1","Q2"], {"收入": [100,200]})
helper.add_summary_slide("总结", ["要点A", "要点B"])
helper.save("output.pptx")
```

### 使用 quick_create 快捷函数

```python
from ppt_helpers import quick_create

quick_create("我的演示", [
    {"type": "content", "title": "要点", "bullets": ["A", "B"]},
    {"type": "chart", "title": "数据", "chart_type": "bar",
     "categories": ["Q1","Q2"], "series_data": {"收入": [100,200]}},
], theme="tech_dark", output="my_ppt.pptx")
```

### 使用 create_ppt.py 命令行工具

```bash
python scripts/create_ppt.py '<json_config>'
python scripts/create_ppt.py --demo
```

### 读取已有 PPT

```bash
python scripts/read_ppt.py existing.pptx
```

## 图表创建参考

详细图表代码示例见 `resources/chart-guide.md`，支持柱状图、饼图、折线图等。

## 模板使用

模板文件存放在 `resources/templates/` 目录下。详细使用指南见 `resources/templates/README.md`。

可用模板：
- `blank.pptx` — 默认空白模板，含基本母版布局
- `company-1.pptx` — 企业模板 1
- `company-2026.pptx` — 企业模板 2026
