@echo off
REM Go-Word 构建脚本 (Windows版本)
REM 支持选择性构建不同的包

setlocal enabledelayedexpansion

echo Go-Word 构建脚本 (Windows)
echo ================================

REM 默认构建所有核心包
set BUILD_CORE=true
set BUILD_GUI=false
set BUILD_EXAMPLES=false
set BUILD_DEMOS=false

REM 解析命令行参数
:parse_args
if "%~1"=="" goto :end_parse
if "%~1"=="--core-only" (
    set BUILD_CORE=true
    set BUILD_GUI=false
    set BUILD_EXAMPLES=false
    set BUILD_DEMOS=false
    shift
    goto :parse_args
)
if "%~1"=="--with-gui" (
    set BUILD_GUI=true
    shift
    goto :parse_args
)
if "%~1"=="--with-examples" (
    set BUILD_EXAMPLES=true
    shift
    goto :parse_args
)
if "%~1"=="--with-demos" (
    set BUILD_DEMOS=true
    shift
    goto :parse_args
)
if "%~1"=="--all" (
    set BUILD_CORE=true
    set BUILD_GUI=true
    set BUILD_EXAMPLES=true
    set BUILD_DEMOS=true
    shift
    goto :parse_args
)
if "%~1"=="--help" (
    echo 用法: %0 [选项]
    echo 选项:
    echo   --core-only     只构建核心包 (默认)
    echo   --with-gui      包含GUI包构建
    echo   --with-examples 包含示例构建
    echo   --with-demos    包含演示构建
    echo   --all           构建所有包
    echo   --help          显示此帮助信息
    exit /b 0
)
echo 未知选项: %~1
echo 使用 --help 查看帮助信息
exit /b 1

:end_parse

echo 构建配置:
echo   核心包: %BUILD_CORE%
echo   GUI包: %BUILD_GUI%
echo   示例: %BUILD_EXAMPLES%
echo   演示: %BUILD_DEMOS%
echo.

REM 清理之前的构建
echo 清理之前的构建...
if exist go-word.exe del go-word.exe
if exist test_package.exe del test_package.exe
if exist import_test.exe del import_test.exe
if exist basic_usage.exe del basic_usage.exe

REM 更新依赖
echo 更新Go模块依赖...
go mod tidy
if errorlevel 1 (
    echo 错误: 更新依赖失败
    exit /b 1
)

REM 构建核心包
if "%BUILD_CORE%"=="true" (
    echo 构建核心包...
    go build -v ./pkg/word/... ./pkg/opc/... ./pkg/parser/... ./pkg/types/... ./pkg/utils/... ./pkg/writer/... ./pkg/plugin/...
    if errorlevel 1 (
        echo 错误: 核心包构建失败
        exit /b 1
    )
    echo ✓ 核心包构建完成
)

REM 构建GUI包
if "%BUILD_GUI%"=="true" (
    echo 构建GUI包...
    go build -v ./pkg/gui/...
    if errorlevel 1 (
        echo ⚠ GUI包构建失败
    ) else (
        echo ✓ GUI包构建完成
    )
)

REM 构建示例
if "%BUILD_EXAMPLES%"=="true" (
    echo 构建示例...
    if exist examples (
        for /d %%i in (examples\*) do (
            if exist "%%i\main.go" (
                echo   构建示例: %%~ni
                go build -o "%%~ni.exe" "%%i"
                if errorlevel 1 (
                    echo     ⚠ 构建失败
                ) else (
                    echo     ✓ 构建成功
                )
            )
        )
    )
    echo ✓ 示例构建完成
)

REM 构建演示
if "%BUILD_DEMOS%"=="true" (
    echo 构建演示...
    if exist demos (
        for /d %%i in (demos\*) do (
            for %%j in ("%%i\*_demo.go") do (
                if exist "%%j" (
                    echo   构建演示: %%~ni
                    go build -o "%%~ni_demo.exe" "%%j"
                    if errorlevel 1 (
                        echo     ⚠ 构建失败
                    ) else (
                        echo     ✓ 构建成功
                    )
                )
            )
        )
    )
    echo ✓ 演示构建完成
)

REM 运行测试
echo 运行核心包测试...
go test ./pkg/word/... -v
if errorlevel 1 (
    echo ⚠ 测试失败
) else (
    echo ✓ 测试通过
)

echo.
echo 🎉 构建完成!
echo.
echo 可执行文件:
dir *.exe 2>nul | findstr /i "\.exe$" >nul && (
    dir *.exe
) || (
    echo   无可执行文件生成
)
echo.
echo 提示: 在Windows环境中，所有包都应该能正常构建
