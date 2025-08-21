#!/bin/bash

# Go-Word 构建脚本
# 支持选择性构建不同的包

set -e

echo "Go-Word 构建脚本"
echo "=================="

# 默认构建所有核心包
BUILD_CORE=true
BUILD_GUI=false
BUILD_EXAMPLES=false
BUILD_DEMOS=false

# 解析命令行参数
while [[ $# -gt 0 ]]; do
    case $1 in
        --core-only)
            BUILD_CORE=true
            BUILD_GUI=false
            BUILD_EXAMPLES=false
            BUILD_DEMOS=false
            shift
            ;;
        --with-gui)
            BUILD_GUI=true
            shift
            ;;
        --with-examples)
            BUILD_EXAMPLES=true
            shift
            ;;
        --with-demos)
            BUILD_DEMOS=true
            shift
            ;;
        --all)
            BUILD_CORE=true
            BUILD_GUI=true
            BUILD_EXAMPLES=true
            BUILD_DEMOS=true
            shift
            ;;
        --help)
            echo "用法: $0 [选项]"
            echo "选项:"
            echo "  --core-only     只构建核心包 (默认)"
            echo "  --with-gui      包含GUI包构建"
            echo "  --with-examples 包含示例构建"
            echo "  --with-demos    包含演示构建"
            echo "  --all           构建所有包"
            echo "  --help          显示此帮助信息"
            exit 0
            ;;
        *)
            echo "未知选项: $1"
            echo "使用 --help 查看帮助信息"
            exit 1
            ;;
    esac
done

echo "构建配置:"
echo "  核心包: $BUILD_CORE"
echo "  GUI包: $BUILD_GUI"
echo "  示例: $BUILD_EXAMPLES"
echo "  演示: $BUILD_DEMOS"
echo ""

# 清理之前的构建
echo "清理之前的构建..."
rm -f go-word
rm -f test_package
rm -f import_test
rm -f basic_usage

# 更新依赖
echo "更新Go模块依赖..."
go mod tidy

# 构建核心包
if [ "$BUILD_CORE" = true ]; then
    echo "构建核心包..."
    go build -v ./pkg/word/... ./pkg/opc/... ./pkg/parser/... ./pkg/types/... ./pkg/utils/... ./pkg/writer/... ./pkg/plugin/...
    echo "✓ 核心包构建完成"
fi

# 构建GUI包 (如果支持)
if [ "$BUILD_GUI" = true ]; then
    echo "构建GUI包..."
    if go build -v ./pkg/gui/... 2>/dev/null; then
        echo "✓ GUI包构建完成"
    else
        echo "⚠ GUI包构建失败 (可能需要X11库或OpenGL支持)"
        echo "  在WSL环境中，可能需要安装X11开发库:"
        echo "  sudo apt-get install libx11-dev libxrandr-dev libxinerama-dev libxcursor-dev libxi-dev"
    fi
fi

# 构建示例
if [ "$BUILD_EXAMPLES" = true ]; then
    echo "构建示例..."
    if [ -d "examples" ]; then
        for example in examples/*/; do
            if [ -d "$example" ] && [ -f "$example/main.go" ]; then
                echo "  构建示例: $(basename "$example")"
                if go build -o "$(basename "$example")" "$example"; then
                    echo "    ✓ 构建成功"
                else
                    echo "    ⚠ 构建失败"
                fi
            fi
        done
    fi
    echo "✓ 示例构建完成"
fi

# 构建演示
if [ "$BUILD_DEMOS" = true ]; then
    echo "构建演示..."
    if [ -d "demos" ]; then
        for demo in demos/*/; do
            if [ -d "$demo" ] && [ -f "$demo"*_demo.go ]; then
                echo "  构建演示: $(basename "$demo")"
                if go build -o "$(basename "$demo")_demo" "$demo"*_demo.go; then
                    echo "    ✓ 构建成功"
                else
                    echo "    ⚠ 构建失败"
                fi
            fi
        done
    fi
    echo "✓ 演示构建完成"
fi

# 运行测试
echo "运行核心包测试..."
go test ./pkg/word/... -v

echo ""
echo "🎉 构建完成!"
echo ""
echo "可执行文件:"
ls -la go-word test_package import_test basic_usage 2>/dev/null || echo "  无可执行文件生成"
echo ""
echo "提示: 在WSL环境中，GUI包可能需要额外的X11库支持"
