# 专业配色方案库

> 6 套配色方案的完整色值（Hex + RGB）、字体配置和使用示例。

---

## 1. 商务蓝 `business_blue`

适用：企业汇报、年报、正式场合

| 角色 | Hex | RGB | 用途 |
|------|-----|-----|------|
| 主色 | `#1B3A5C` | `RGBColor(0x1B, 0x3A, 0x5C)` | 标题栏、封面背景 |
| 辅色 | `#4A90D9` | `RGBColor(0x4A, 0x90, 0xD9)` | 章节页、图表 |
| 强调色 | `#F39C12` | `RGBColor(0xF3, 0x9C, 0x12)` | 高亮、总结页 |
| 背景色 | `#FFFFFF` | `RGBColor(0xFF, 0xFF, 0xFF)` | 内容页背景 |
| 文字色 | `#2C3E50` | `RGBColor(0x2C, 0x3E, 0x50)` | 正文文字 |
| 浅背景 | `#ECF0F1` | `RGBColor(0xEC, 0xF0, 0xF1)` | 辅助色块 |

- **中文字体**: Microsoft YaHei（微软雅黑）
- **英文字体**: Calibri

```python
helper = PPTHelper(theme="business_blue")
```

---

## 2. 科技感 `tech_dark`

适用：技术分享、产品发布、科技主题

| 角色 | Hex | RGB | 用途 |
|------|-----|-----|------|
| 主色 | `#0D1B2A` | `RGBColor(0x0D, 0x1B, 0x2A)` | 深色背景 |
| 辅色 | `#1B998B` | `RGBColor(0x1B, 0x99, 0x8B)` | 章节页 |
| 强调色 | `#00F5D4` | `RGBColor(0x00, 0xF5, 0xD4)` | 荧光高亮 |
| 背景色 | `#0D1B2A` | `RGBColor(0x0D, 0x1B, 0x2A)` | 深色背景 |
| 文字色 | `#E0E0E0` | `RGBColor(0xE0, 0xE0, 0xE0)` | 浅色正文 |
| 浅背景 | `#1B2A3A` | `RGBColor(0x1B, 0x2A, 0x3A)` | 次级色块 |

- **中文字体**: Source Han Sans CN（思源黑体）
- **英文字体**: Consolas

```python
helper = PPTHelper(theme="tech_dark")
```

---

## 3. 清新绿 `fresh_green`

适用：教育、健康、环保主题

| 角色 | Hex | RGB | 用途 |
|------|-----|-----|------|
| 主色 | `#2D6A4F` | `RGBColor(0x2D, 0x6A, 0x4F)` | 标题栏 |
| 辅色 | `#74C69D` | `RGBColor(0x74, 0xC6, 0x9D)` | 章节页 |
| 强调色 | `#40916C` | `RGBColor(0x40, 0x91, 0x6C)` | 高亮 |
| 背景色 | `#FFFFFF` | `RGBColor(0xFF, 0xFF, 0xFF)` | 白色背景 |
| 文字色 | `#2D3E2D` | `RGBColor(0x2D, 0x3E, 0x2D)` | 深绿文字 |
| 浅背景 | `#D8F3DC` | `RGBColor(0xD8, 0xF3, 0xDC)` | 浅绿色块 |

- **中文字体**: Microsoft YaHei（微软雅黑）
- **英文字体**: Segoe UI

```python
helper = PPTHelper(theme="fresh_green")
```

---

## 4. 暖橙 `warm_orange`

适用：营销、创意提案、品牌推广

| 角色 | Hex | RGB | 用途 |
|------|-----|-----|------|
| 主色 | `#E76F51` | `RGBColor(0xE7, 0x6F, 0x51)` | 标题栏 |
| 辅色 | `#F4A261` | `RGBColor(0xF4, 0xA2, 0x61)` | 章节页 |
| 强调色 | `#264653` | `RGBColor(0x26, 0x46, 0x53)` | 对比强调 |
| 背景色 | `#FFFFFF` | `RGBColor(0xFF, 0xFF, 0xFF)` | 白色背景 |
| 文字色 | `#264653` | `RGBColor(0x26, 0x46, 0x53)` | 深色文字 |
| 浅背景 | `#FDF0E2` | `RGBColor(0xFD, 0xF0, 0xE2)` | 暖色块 |

- **中文字体**: FZLanTingHei-R-GBK（方正兰亭黑）
- **英文字体**: Arial

```python
helper = PPTHelper(theme="warm_orange")
```

---

## 5. 优雅紫 `elegant_purple`

适用：设计、文化、艺术、高端品牌

| 角色 | Hex | RGB | 用途 |
|------|-----|-----|------|
| 主色 | `#432371` | `RGBColor(0x43, 0x23, 0x71)` | 标题栏 |
| 辅色 | `#714EBF` | `RGBColor(0x71, 0x4E, 0xBF)` | 章节页 |
| 强调色 | `#FAAE7B` | `RGBColor(0xFA, 0xAE, 0x7B)` | 暖色高亮 |
| 背景色 | `#FFFFFF` | `RGBColor(0xFF, 0xFF, 0xFF)` | 白色背景 |
| 文字色 | `#331A55` | `RGBColor(0x33, 0x1A, 0x55)` | 深紫文字 |
| 浅背景 | `#F0E6FA` | `RGBColor(0xF0, 0xE6, 0xFA)` | 浅紫色块 |

- **中文字体**: Source Han Serif CN（思源宋体）
- **英文字体**: Georgia

```python
helper = PPTHelper(theme="elegant_purple")
```

---

## 6. 简约灰 `minimal_gray`

适用：通用场景、极简风格

| 角色 | Hex | RGB | 用途 |
|------|-----|-----|------|
| 主色 | `#2B2D42` | `RGBColor(0x2B, 0x2D, 0x42)` | 标题栏 |
| 辅色 | `#8D99AE` | `RGBColor(0x8D, 0x99, 0xAE)` | 章节页 |
| 强调色 | `#EF233C` | `RGBColor(0xEF, 0x23, 0x3C)` | 红色高亮 |
| 背景色 | `#FFFFFF` | `RGBColor(0xFF, 0xFF, 0xFF)` | 白色背景 |
| 文字色 | `#2B2D42` | `RGBColor(0x2B, 0x2D, 0x42)` | 深色文字 |
| 浅背景 | `#EDF2F4` | `RGBColor(0xED, 0xF2, 0xF4)` | 浅灰色块 |

- **中文字体**: Microsoft YaHei（微软雅黑）
- **英文字体**: Helvetica

```python
helper = PPTHelper(theme="minimal_gray")
```
