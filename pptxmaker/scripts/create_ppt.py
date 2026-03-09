"""
通用 PPT 创建脚本 - 接收 JSON 参数
====================================
用法:
    python create_ppt.py '<json_config>'

JSON 参数示例:
{
    "title": "演示标题",
    "theme": "business_blue",
    "template": null,
    "output": "output.pptx",
    "slides": [
        {"type": "content", "title": "第一页", "bullets": ["要点1", "要点2"]},
        {"type": "chart", "title": "数据", "chart_type": "bar",
         "categories": ["Q1","Q2"], "series_data": {"收入": [100,200]}}
    ]
}
"""

import sys
import json
import os

# 确保可以导入同目录下的 ppt_helpers
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ppt_helpers import PPTHelper, quick_create


def main():
    if len(sys.argv) < 2:
        print("用法: python create_ppt.py '<json_config>'")
        print("  或: python create_ppt.py --demo  (生成演示文件)")
        sys.exit(1)

    if sys.argv[1] == "--demo":
        # 演示模式：生成一个示例 PPT
        output = quick_create(
            title="PPT Maker 演示",
            theme="business_blue",
            output="demo_presentation.pptx",
            slides_config=[
                {
                    "type": "content",
                    "title": "功能特性",
                    "bullets": [
                        "支持 6 种专业配色方案",
                        "内置 8 种页面布局模式",
                        "原生 PowerPoint 图表支持",
                        "企业模板加载与继承",
                        "中英文字体自动适配",
                    ],
                },
                {
                    "type": "two_column",
                    "title": "方案对比",
                    "left_title": "优势",
                    "left_items": ["输出完全可编辑", "依赖极轻", "原生图表"],
                    "right_title": "适用场景",
                    "right_items": ["企业汇报", "技术分享", "产品发布"],
                },
                {
                    "type": "chart",
                    "title": "季度数据",
                    "chart_type": "bar",
                    "categories": ["Q1", "Q2", "Q3", "Q4"],
                    "series_data": {"收入": [120, 180, 240, 300]},
                },
                {
                    "type": "summary",
                    "title": "总结",
                    "key_points": [
                        "python-pptx 方案综合评分最高",
                        "原生可编辑 PPTX 输出",
                        "插件已通过评审，可投入使用",
                    ],
                },
            ],
        )
        print(f"✅ 演示文件已创建: {output}")
        return

    # 正常模式：解析 JSON 配置
    try:
        config = json.loads(sys.argv[1])
    except json.JSONDecodeError as e:
        print(f"❌ JSON 解析错误: {e}")
        sys.exit(1)

    title = config.get("title", "Untitled Presentation")
    theme = config.get("theme", "business_blue")
    template = config.get("template", None)
    output = config.get("output", "presentation.pptx")
    slides = config.get("slides", [])

    try:
        result = quick_create(
            title=title,
            slides_config=slides,
            theme=theme,
            template=template,
            output=output,
        )
        print(f"✅ PPT 已创建: {result}")
        print(f"📄 幻灯片数: {len(slides) + 1}")  # +1 for title slide
        print(f"🎨 配色方案: {theme}")
    except Exception as e:
        print(f"❌ 创建失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
