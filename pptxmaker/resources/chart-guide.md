# 图表创建指南

> python-pptx 原生图表的完整代码示例，支持柱状图、饼图、折线图。

---

## 通用导入

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
```

---

## 1. 柱状图 (Bar/Column Chart)

### 基础柱状图

```python
chart_data = CategoryChartData()
chart_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
chart_data.add_series('收入 (万元)', (120, 180, 240, 300))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1.5), Inches(11), Inches(5.5),
    chart_data
)
```

### 多系列柱状图

```python
chart_data = CategoryChartData()
chart_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
chart_data.add_series('收入', (120, 180, 240, 300))
chart_data.add_series('支出', (80, 120, 160, 200))
chart_data.add_series('利润', (40, 60, 80, 100))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1.5), Inches(11), Inches(5.5),
    chart_data
)

# 样式定制
chart = chart_frame.chart
chart.has_legend = True
chart.legend.include_in_layout = False
```

### 堆叠柱状图

```python
chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_STACKED,  # 堆叠
    Inches(1), Inches(1.5), Inches(11), Inches(5.5),
    chart_data
)
```

### 条形图（水平）

```python
chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED,  # 水平条形
    Inches(1), Inches(1.5), Inches(11), Inches(5.5),
    chart_data
)
```

---

## 2. 饼图 (Pie Chart)

### 基础饼图

```python
chart_data = CategoryChartData()
chart_data.categories = ['产品A', '产品B', '产品C', '其他']
chart_data.add_series('市场份额', (35, 28, 22, 15))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.PIE,
    Inches(2), Inches(1.5), Inches(9), Inches(5.5),
    chart_data
)

# 显示百分比标签
chart = chart_frame.chart
chart.has_legend = True
plot = chart.plots[0]
plot.has_data_labels = True
data_labels = plot.data_labels
data_labels.number_format = '0%'
data_labels.show_percentage = True
data_labels.show_value = False
```

### 环形图

```python
chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT,  # 环形
    Inches(2), Inches(1.5), Inches(9), Inches(5.5),
    chart_data
)
```

---

## 3. 折线图 (Line Chart)

### 基础折线图

```python
chart_data = CategoryChartData()
chart_data.categories = ['1月', '2月', '3月', '4月', '5月', '6月']
chart_data.add_series('用户数 (万)', (50, 65, 80, 110, 145, 180))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE_MARKERS,
    Inches(1), Inches(1.5), Inches(11), Inches(5.5),
    chart_data
)

chart = chart_frame.chart
chart.has_legend = True
```

### 多系列折线图

```python
chart_data = CategoryChartData()
chart_data.categories = ['1月', '2月', '3月', '4月', '5月', '6月']
chart_data.add_series('移动端', (30, 45, 55, 70, 90, 120))
chart_data.add_series('Web端', (20, 20, 25, 40, 55, 60))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE_MARKERS,
    Inches(1), Inches(1.5), Inches(11), Inches(5.5),
    chart_data
)
```

### 平滑折线（无标记点）

```python
chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE,  # 无标记点
    Inches(1), Inches(1.5), Inches(11), Inches(5.5),
    chart_data
)
```

---

## 使用辅助库

以上图表均可通过 `PPTHelper` 一行代码创建：

```python
helper.add_chart_slide(
    "季度收入对比",
    "bar",                                    # bar / pie / line
    ["Q1", "Q2", "Q3", "Q4"],               # 分类
    {"收入": [120, 180, 240, 300],           # 数据系列
     "支出": [80, 120, 160, 200]}
)
```

支持的 `chart_type` 值：
- `"bar"` → 柱状图 (`COLUMN_CLUSTERED`)
- `"pie"` → 饼图 (`PIE`)
- `"line"` → 折线图 (`LINE_MARKERS`)
