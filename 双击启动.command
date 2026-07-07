#!/bin/bash
# ==========================================
#  优惠券批量生成器 - macOS 启动脚本
#  双击此文件即可启动
# ==========================================

# 清理可能干扰 Python 的 DYLD 环境变量
unset DYLD_LIBRARY_PATH
unset DYLD_FALLBACK_LIBRARY_PATH
unset DYLD_INSERT_LIBRARIES
unset DYLD_FORCE_FLAT_NAMESPACE

# 切换到脚本所在目录
cd "$(dirname "$0")"

# 查找 Python 3
PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" &> /dev/null; then
        version=$("$cmd" -c 'import sys; print(sys.version_info[0])' 2>/dev/null)
        if [ "$version" = "3" ]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo "========================================"
    echo "  错误：未找到 Python 3"
    echo "  Error: Python 3 not found"
    echo "========================================"
    echo ""
    echo "请安装 Python 3.9+："
    echo "  方式一：从 https://www.python.org/downloads/ 下载安装"
    echo "  方式二：使用 Homebrew: brew install python3"
    echo ""
    echo "按任意键退出..."
    read -n 1
    exit 1
fi

# 检查必要依赖
echo "检查依赖..."
$PYTHON -c "import qrcode, PIL, reportlab" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "正在安装依赖包..."
    $PYTHON -m pip install qrcode Pillow reportlab openpyxl xlrd
    if [ $? -ne 0 ]; then
        echo ""
        echo "依赖安装失败，请手动执行："
        echo "  $PYTHON -m pip install qrcode Pillow reportlab openpyxl xlrd"
        echo ""
        echo "按任意键退出..."
        read -n 1
        exit 1
    fi
    echo "依赖安装完成！"
fi

echo ""
echo "=========================================="
echo "  优惠券批量生成器"
echo "  Coupon Batch Generator"
echo "=========================================="
echo ""
echo "正在启动服务... / Starting server..."
echo "浏览器将自动打开 / Browser will open automatically"
echo ""

$PYTHON coupon_web.py

echo ""
echo "服务已停止。按任意键退出..."
read -n 1
