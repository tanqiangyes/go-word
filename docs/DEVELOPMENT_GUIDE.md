# 开发指南

## 概述

本文档为 Go OpenXML SDK 项目的开发者提供详细的开发指南，包括环境设置、项目结构、开发流程和贡献指南。

## 目录

- [环境要求](#环境要求)
- [项目设置](#项目设置)
- [项目结构](#项目结构)
- [开发流程](#开发流程)
- [测试指南](#测试指南)
- [代码规范](#代码规范)
- [贡献指南](#贡献指南)
- [发布流程](#发布流程)

## 环境要求

### 必需软件

- **Go**: 1.18 或更高版本
- **Git**: 最新版本
- **编辑器**: VS Code、GoLand 或 Vim/Emacs

### 推荐工具

- **Docker**: 用于容器化测试
- **Make**: 用于构建脚本
- **golangci-lint**: 代码质量检查

### 安装 Go

#### Windows
```bash
# 下载并安装 Go
# 访问 https://golang.org/dl/
# 下载 Windows 安装包并运行

# 验证安装
go version
```

#### macOS
```bash
# 使用 Homebrew
brew install go

# 验证安装
go version
```

#### Linux
```bash
# Ubuntu/Debian
sudo apt-get update
sudo apt-get install golang-go

# CentOS/RHEL
sudo yum install golang

# 验证安装
go version
```

## 项目设置

### 1. 克隆项目

```bash
git clone https://github.com/tanqiangyes/go-word.git
cd go-word
```

### 2. 安装依赖

```bash
go mod tidy
```

### 3. 验证设置

```bash
# 运行测试
go test ./...

# 运行示例
go run examples/basic_usage.go
```

### 4. 开发环境配置

#### VS Code 配置

创建 `.vscode/settings.json`:

```json
{
    "go.useLanguageServer": true,
    "go.lintTool": "golangci-lint",
    "go.lintFlags": ["--fast"],
    "go.testFlags": ["-v"],
    "go.coverOnSave": true,
    "go.coverOnTestPackage": true,
    "go.coverOnSingleTest": true,
    "go.testOnSave": true,
    "go.buildOnSave": true,
    "go.vetOnSave": true,
    "go.formatTool": "goimports",
    "go.formatFlags": ["-w"],
    "editor.formatOnSave": true,
    "editor.codeActionsOnSave": {
        "source.organizeImports": true
    }
}
```

#### golangci-lint 配置

创建 `.golangci.yml`:

```yaml
run:
  timeout: 5m
  modules-download-mode: readonly

linters:
  enable:
    - gofmt
    - goimports
    - govet
    - errcheck
    - staticcheck
    - gosimple
    - ineffassign
    - unused
    - misspell
    - gosec

linters-settings:
  gofmt:
    simplify: true
  goimports:
    local-prefixes: github.com/tanqiangyes/go-word

issues:
  exclude-rules:
    - path: _test\.go
      linters:
        - errcheck
```

## 项目结构

```
go-word/
├── .cursor/                 # Cursor IDE 配置
│   └── scopes/             # 项目规范文档
├── docs/                   # 文档
│   ├── API_REFERENCE.md    # API 参考
│   ├── DEVELOPMENT_GUIDE.md # 开发指南
│   └── PROJECT_SUMMARY.md  # 项目总结
├── examples/               # 使用示例
│   ├── basic_usage.go      # 基本用法
│   ├── advanced_usage.go   # 高级用法
│   └── ...                 # 更多示例
├── pkg/                    # 核心包
│   ├── opc/               # OPC 容器处理
│   ├── word/  # Word 文档处理
│   ├── parser/            # XML 解析器
│   ├── writer/            # 文档写入器
│   ├── types/             # 共享类型定义
│   └── utils/             # 工具函数
├── tests/                 # 测试文件
├── .gitignore            # Git 忽略文件
├── go.mod                # Go 模块定义
├── go.sum                # 依赖校验和
├── README.md             # 项目说明
└── TODO.md              # 待办事项
```

### 包说明

#### pkg/opc
OPC（Open Packaging Convention）容器处理，负责 Word 文档的 ZIP 容器操作。

**主要功能:**
- 打开/关闭 Word 文档
- 访问文档内部文件
- 管理文档部分

#### pkg/word
Word 文档处理核心包，提供文档的读取、解析和操作功能。

**主要功能:**
- 文档内容提取
- 段落和表格处理
- 高级格式化
- 文档保护

#### pkg/parser
XML 解析器，专门处理 word XML 格式。

**主要功能:**
- WordML XML 解析
- 通用 XML 解析
- 命名空间处理

#### pkg/writer
文档写入器，提供文档创建和修改功能。

**主要功能:**
- 创建新文档
- 添加段落和表格
- 保存文档

#### pkg/types
共享类型定义，解决包间导入循环问题。

**主要类型:**
- Document、Paragraph、Run
- Table、TableRow、TableCell
- 各种格式化结构

#### pkg/utils
工具函数包，提供错误处理和通用工具。

**主要功能:**
- 错误处理
- 日志记录
- 通用工具函数

## 开发流程

### 1. 功能开发流程

#### 步骤 1: 创建功能分支

```bash
# 确保在主分支
git checkout main
git pull origin main

# 创建功能分支
git checkout -b feature/your-feature-name
```

#### 步骤 2: 开发功能

1. **编写代码**: 在相应的包中实现功能
2. **添加测试**: 为新功能编写单元测试
3. **更新文档**: 更新 API 文档和示例
4. **运行测试**: 确保所有测试通过

```bash
# 运行测试
go test ./...

# 运行覆盖率测试
go test -cover ./pkg/...

# 代码质量检查
golangci-lint run
```

#### 步骤 3: 提交代码

```bash
# 添加文件
git add .

# 提交代码
git commit -m "feat: add new feature description"

# 推送到远程
git push origin feature/your-feature-name
```

#### 步骤 4: 创建 Pull Request

1. 在 GitHub 上创建 Pull Request
2. 填写详细的描述
3. 等待代码审查
4. 根据反馈修改代码

### 2. Bug 修复流程

#### 步骤 1: 创建 Bug 分支

```bash
git checkout -b fix/bug-description
```

#### 步骤 2: 修复 Bug

1. **重现问题**: 编写测试用例重现 Bug
2. **修复代码**: 修复问题
3. **验证修复**: 运行测试确保修复有效
4. **添加测试**: 添加回归测试

#### 步骤 3: 提交修复

```bash
git commit -m "fix: description of the bug fix"
git push origin fix/bug-description
```

### 3. 代码审查流程

#### 审查清单

- [ ] 代码符合项目规范
- [ ] 所有测试通过
- [ ] 测试覆盖率不降低
- [ ] 文档已更新
- [ ] 示例代码已更新
- [ ] 没有引入新的依赖
- [ ] 性能影响已评估

#### 审查要点

1. **功能正确性**: 代码是否实现了预期功能
2. **代码质量**: 代码是否清晰、可维护
3. **错误处理**: 是否正确处理了错误情况
4. **性能影响**: 是否对性能有负面影响
5. **安全性**: 是否引入了安全风险

## 测试指南

### 测试类型

#### 1. 单元测试

**位置**: 与源代码文件同目录的 `*_test.go` 文件

**命名规范**:
```go
func TestFunctionName(t *testing.T) {
    // 测试代码
}

func TestFunctionName_Scenario(t *testing.T) {
    // 特定场景测试
}
```

**示例**:
```go
func TestDocumentOpen(t *testing.T) {
    doc, err := word.Open("testdata/sample.docx")
    if err != nil {
        t.Fatalf("Failed to open document: %v", err)
    }
    defer doc.Close()
    
    // 验证文档内容
    text, err := doc.GetText()
    if err != nil {
        t.Fatalf("Failed to get text: %v", err)
    }
    
    if text == "" {
        t.Error("Expected non-empty text content")
    }
}
```

#### 2. 集成测试

**位置**: `tests/` 目录

**目的**: 测试多个组件的交互

**示例**:
```go
func TestDocumentCreationAndModification(t *testing.T) {
    // 创建文档
    docWriter := writer.NewDocumentWriter()
    err := docWriter.CreateNewDocument()
    if err != nil {
        t.Fatalf("Failed to create document: %v", err)
    }
    
    // 添加内容
    err = docWriter.AddParagraph("Test paragraph", "Normal")
    if err != nil {
        t.Fatalf("Failed to add paragraph: %v", err)
    }
    
    // 保存文档
    tempFile := "temp_test_document.docx"
    err = docWriter.Save(tempFile)
    if err != nil {
        t.Fatalf("Failed to save document: %v", err)
    }
    defer os.Remove(tempFile)
    
    // 验证保存的文档
    doc, err := word.Open(tempFile)
    if err != nil {
        t.Fatalf("Failed to open saved document: %v", err)
    }
    defer doc.Close()
    
    text, err := doc.GetText()
    if err != nil {
        t.Fatalf("Failed to get text from saved document: %v", err)
    }
    
    if !strings.Contains(text, "Test paragraph") {
        t.Error("Saved document does not contain expected text")
    }
}
```

#### 3. 性能基准测试

**位置**: `tests/` 目录

**命名规范**: `BenchmarkFunctionName`

**示例**:
```go
func BenchmarkDocumentOpen(b *testing.B) {
    for i := 0; i < b.N; i++ {
        doc, err := word.Open("testdata/large_document.docx")
        if err != nil {
            b.Fatalf("Failed to open document: %v", err)
        }
        doc.Close()
    }
}

func BenchmarkTextExtraction(b *testing.B) {
    doc, err := word.Open("testdata/large_document.docx")
    if err != nil {
        b.Fatalf("Failed to open document: %v", err)
    }
    defer doc.Close()
    
    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        _, err := doc.GetText()
        if err != nil {
            b.Fatalf("Failed to get text: %v", err)
        }
    }
}
```

### 测试运行

#### 运行所有测试
```bash
go test ./...
```

#### 运行特定包的测试
```bash
go test ./pkg/word
```

#### 运行覆盖率测试
```bash
go test -cover ./pkg/...
```

#### 生成覆盖率报告
```bash
go test -coverprofile=coverage.out ./pkg/...
go tool cover -html=coverage.out -o coverage.html
```

#### 运行基准测试
```bash
go test -bench=. ./tests/
```

### 测试数据

#### 测试文档

在 `testdata/` 目录中放置测试用的 Word 文档：

```
testdata/
├── sample.docx          # 基本测试文档
├── large_document.docx  # 大文档性能测试
├── complex_format.docx  # 复杂格式测试
└── invalid_format.docx  # 无效格式测试
```

#### 测试文档要求

1. **sample.docx**: 包含基本段落、表格和格式化
2. **large_document.docx**: 大文档用于性能测试
3. **complex_format.docx**: 包含复杂格式的文档
4. **invalid_format.docx**: 故意损坏的文档用于错误测试

## 代码规范

### 1. 命名规范

#### 包名
- 使用小写字母
- 避免下划线
- 简洁明了

```go
// 正确
package word
package opc
package parser

// 错误
package word
package word_processing_ml
```

#### 函数名
- 使用驼峰命名法
- 动词开头
- 描述性名称

```go
// 正确
func OpenDocument(filename string) (*Document, error)
func GetText() (string, error)
func AddParagraph(text string) error

// 错误
func open_document(filename string) (*Document, error)
func get_text() (string, error)
```

#### 变量名
- 使用驼峰命名法
- 描述性名称
- 避免缩写

```go
// 正确
var documentContent *DocumentContent
var paragraphCount int
var isFormatted bool

// 错误
var doc *DocumentContent
var pCount int
var fmt bool
```

#### 常量名
- 使用大写字母和下划线
- 描述性名称

```go
// 正确
const (
    DefaultFontSize = 12
    MaxTableRows    = 1000
    WordMLNamespace = "http://schemas.openxmlformats.org/word/2006/main"
)

// 错误
const (
    defaultFontSize = 12
    maxTableRows    = 1000
)
```

### 2. 代码格式

#### 导入顺序
```go
import (
    // 标准库
    "fmt"
    "io"
    "os"
    
    // 第三方库
    "github.com/example/package"
    
    // 本地包
    "github.com/tanqiangyes/go-word/pkg/types"
)
```

#### 函数结构
```go
func FunctionName(param1 string, param2 int) (result string, err error) {
    // 参数验证
    if param1 == "" {
        return "", fmt.Errorf("param1 cannot be empty")
    }
    
    // 主要逻辑
    result = processData(param1, param2)
    
    // 返回结果
    return result, nil
}
```

#### 错误处理
```go
// 使用 fmt.Errorf 包装错误
if err != nil {
    return fmt.Errorf("failed to process document: %w", err)
}

// 使用 errors.Is 检查错误类型
if errors.Is(err, &DocumentError{}) {
    // 处理文档错误
}
```

### 3. 注释规范

#### 包注释
```go
// Package word provides word document processing functionality.
// It includes functions for reading, parsing, and manipulating Word documents.
package word
```

#### 函数注释
```go
// Open opens a Word document from a file.
// It returns a Document object that can be used to access the document's content.
// The returned document must be closed when no longer needed.
func Open(filename string) (*Document, error) {
    // 实现代码
}
```

#### 类型注释
```go
// Document represents a Word document.
// It contains the document's content, metadata, and provides methods for accessing
// and manipulating the document.
type Document struct {
    container *opc.Container
    mainPart  *MainDocumentPart
    parts     map[string]*opc.Part
}
```

#### 内联注释
```go
// 处理文档内容
content, err := parseDocumentContent(data)
if err != nil {
    return nil, fmt.Errorf("failed to parse content: %w", err)
}

// 验证解析结果
if content == nil {
    return nil, fmt.Errorf("parsed content is nil")
}
```

### 4. 错误处理

#### 错误类型定义
```go
// DocumentError represents an error related to document operations
type DocumentError struct {
    Message string
    Cause   error
}

func (e *DocumentError) Error() string {
    if e.Cause != nil {
        return fmt.Sprintf("document error: %s: %v", e.Message, e.Cause)
    }
    return fmt.Sprintf("document error: %s", e.Message)
}

func (e *DocumentError) Unwrap() error {
    return e.Cause
}
```

#### 错误创建
```go
// 创建新错误
return &DocumentError{
    Message: "invalid document format",
    Cause:   err,
}

// 使用工具函数
return utils.NewError("invalid document format")
```

### 5. 性能考虑

#### 内存管理
```go
// 使用 defer 确保资源释放
func ProcessDocument(filename string) error {
    doc, err := word.Open(filename)
    if err != nil {
        return err
    }
    defer doc.Close() // 确保文档被关闭
    
    // 处理文档...
    return nil
}
```

#### 避免内存泄漏
```go
// 处理大量文档时，及时释放资源
for _, filename := range filenames {
    doc, err := word.Open(filename)
    if err != nil {
        continue
    }
    
    // 处理文档
    processDocument(doc)
    
    // 立即关闭释放内存
    doc.Close()
}
```

## 贡献指南

### 1. 贡献类型

#### 功能开发
- 新功能实现
- 性能优化
- 架构改进

#### Bug 修复
- 错误修复
- 兼容性改进
- 安全漏洞修复

#### 文档改进
- API 文档更新
- 示例代码改进
- 教程编写

#### 测试改进
- 测试覆盖率提升
- 新测试用例
- 性能基准测试

### 2. 贡献流程

#### 步骤 1: 创建 Issue
1. 在 GitHub 上创建 Issue
2. 描述问题或功能需求
3. 提供复现步骤（如果是 Bug）
4. 讨论解决方案

#### 步骤 2: Fork 项目
1. Fork 项目到自己的账户
2. 克隆到本地
3. 添加上游仓库

```bash
git clone https://github.com/your-username/go-word.git
cd go-word
git remote add upstream https://github.com/tanqiangyes/go-word.git
```

#### 步骤 3: 创建分支
```bash
git checkout -b feature/your-feature-name
# 或
git checkout -b fix/bug-description
```

#### 步骤 4: 开发功能
1. 编写代码
2. 添加测试
3. 更新文档
4. 运行测试

```bash
# 运行测试
go test ./...

# 代码质量检查
golangci-lint run

# 格式化代码
go fmt ./...
```

#### 步骤 5: 提交代码
```bash
git add .
git commit -m "feat: add new feature description"
git push origin feature/your-feature-name
```

#### 步骤 6: 创建 Pull Request
1. 在 GitHub 上创建 Pull Request
2. 填写详细的描述
3. 链接相关的 Issue
4. 等待代码审查

### 3. 提交信息规范

#### 格式
```
<type>(<scope>): <description>

[optional body]

[optional footer]
```

#### 类型
- `feat`: 新功能
- `fix`: Bug 修复
- `docs`: 文档更新
- `style`: 代码格式修改
- `refactor`: 代码重构
- `test`: 测试相关
- `chore`: 构建过程或辅助工具的变动

#### 示例
```
feat(word): add support for table cell merging

- Add MergeCells function to AdvancedFormatter
- Support cell reference parsing (A1, B2, etc.)
- Add comprehensive tests for merge functionality

Closes #123
```

### 4. 代码审查

#### 审查要点
1. **功能正确性**: 代码是否实现了预期功能
2. **代码质量**: 代码是否清晰、可维护
3. **测试覆盖**: 是否有足够的测试
4. **性能影响**: 是否对性能有负面影响
5. **文档更新**: 是否更新了相关文档

#### 审查流程
1. 自动检查（CI/CD）
2. 代码审查者审查
3. 根据反馈修改
4. 最终合并

## 发布流程

### 1. 版本管理

#### 版本号规范
使用语义化版本控制（Semantic Versioning）：

- `MAJOR.MINOR.PATCH`
- `MAJOR`: 不兼容的 API 修改
- `MINOR`: 向下兼容的功能性新增
- `PATCH`: 向下兼容的问题修正

#### 版本标签
```bash
# 创建版本标签
git tag -a v1.0.0 -m "Release version 1.0.0"
git push origin v1.0.0
```

### 2. 发布检查清单

#### 代码质量
- [ ] 所有测试通过
- [ ] 代码质量检查通过
- [ ] 没有已知的 Bug
- [ ] 性能基准测试通过

#### 文档
- [ ] README.md 已更新
- [ ] API 文档已更新
- [ ] 示例代码已更新
- [ ] CHANGELOG.md 已更新

#### 发布准备
- [ ] 版本号已更新
- [ ] 依赖已更新
- [ ] 发布说明已准备
- [ ] GitHub Release 已创建

### 3. 发布步骤

#### 步骤 1: 准备发布
```bash
# 确保在主分支
git checkout main
git pull origin main

# 更新版本号
# 编辑 go.mod 文件
```

#### 步骤 2: 运行测试
```bash
# 运行所有测试
go test ./...

# 运行覆盖率测试
go test -cover ./pkg/...

# 运行基准测试
go test -bench=. ./tests/
```

#### 步骤 3: 创建发布
```bash
# 提交版本更新
git add .
git commit -m "chore: prepare release v1.0.0"

# 创建标签
git tag -a v1.0.0 -m "Release version 1.0.0"

# 推送标签
git push origin v1.0.0
```

#### 步骤 4: 创建 GitHub Release
1. 在 GitHub 上创建 Release
2. 填写发布说明
3. 上传构建产物（如果需要）

### 4. 发布后维护

#### 监控
- 监控 Issue 和 Pull Request
- 关注用户反馈
- 跟踪性能指标

#### 维护
- 及时修复 Bug
- 更新依赖
- 改进文档

---

**注意**: 本指南会随着项目发展而更新，请定期查看最新版本。 