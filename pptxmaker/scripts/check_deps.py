"""SessionStart hook - 检查 python-pptx 依赖是否安装"""
import sys

try:
    import pptx
    # 依赖已安装，静默退出
    sys.exit(0)
except ImportError:
    # 输出提示（会作为 hook 消息显示给用户）
    print("⚠️ python-pptx is not installed. Run /ppt-maker:setup-deps to install.")
    sys.exit(0)
