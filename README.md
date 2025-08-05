# Go OpenXML SDK

一个用 Go 语言重写的 Microsoft Open XML SDK，专门用于 Word 文档处理。提供高性能、低内存占用的 Word 文档解析和操作功能。

## 🚀 特性

- **高性能解析**: 优化的 XML 解析速度，最小化内存占用
- **完整格式支持**: 支持 Word 文档格式（.docx），包含新老版本
- **Go 原生设计**: 遵循 Go 语言最佳实践和惯用法
- **类型安全**: 强类型系统，编译时错误检查
- **详细文档**: 完整的 API 文档和使用示例
- **错误处理**: 结构化的错误信息和日志记录
- **测试覆盖**: 全面的单元测试和性能基准

## 📦 安装

### 使用 Go Modules

```bash
go get github.com/tanqiangyes/go-word
```

### 手动安装

```bash
git clone https://github.com/tanqiangyes/go-word.git
cd go-word
go mod tidy
```

## 🚀 快速开始

### 基本使用

#### 读取 Word 文档

```go
package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func main() {
	// 打开 Word 文档
	doc, err := wordprocessingml.Open("document.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()

	// 获取文档内容
	content, err := doc.GetText()
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("文档内容:", content)

	// 获取段落
	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		log.Fatal(err)
	}

	for i, paragraph := range paragraphs {
		fmt.Printf("段落 %d: %s\n", i+1, paragraph.Text)
		for j, run := range paragraph.Runs {
			fmt.Printf("  运行 %d: '%s' (粗体: %v, 斜体: %v)\n",
				j+1, run.Text, run.Bold, run.Italic)
		}
	}

	// 获取表格
	tables, err := doc.GetTables()
	if err != nil {
		log.Fatal(err)
	}

	for i, table := range tables {
		fmt.Printf("表格 %d: %d行 x %d列\n", i+1, len(table.Rows), table.Columns)
	}
}
```

#### 创建和修改文档

```go
package main

import (
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
	"github.com/tanqiangyes/go-word/pkg/types"
)

func main() {
	// 创建新文档
	docWriter := writer.NewDocumentWriter()
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal(err)
	}

	// 添加段落
	err = docWriter.AddParagraph("这是一个新文档", "Normal")
	if err != nil {
		log.Fatal(err)
	}

	// 添加带格式的段落
	formattedRuns := []types.Run{
		{
			Text:     "粗体文本",
			Bold:     true,
			FontSize: 16,
		},
		{
			Text:     "斜体文本",
			Italic:   true,
			FontSize: 14,
		},
	}
	err = docWriter.AddFormattedParagraph("格式化段落", "Normal", formattedRuns)
	if err != nil {
		log.Fatal(err)
	}

	// 添加表格
	tableData := [][]string{
		{"姓名", "年龄", "职业"},
		{"张三", "25", "工程师"},
		{"李四", "30", "设计师"},
	}
	err = docWriter.AddTable(tableData)
	if err != nil {
		log.Fatal(err)
	}

	// 保存文档
	err = docWriter.Save("new_document.docx")
	if err != nil {
		log.Fatal(err)
	}
}
```

## 📚 API 概览

### 核心包

#### `pkg/wordprocessingml` - Word 文档处理
- `Open(filename string) (*Document, error)` - 打开 Word 文档
- `GetText() (string, error)` - 获取纯文本内容
- `GetParagraphs() ([]Paragraph, error)` - 获取所有段落
- `GetTables() ([]Table, error)` - 获取所有表格

#### `pkg/writer` - 文档写入器
- `NewDocumentWriter() *DocumentWriter` - 创建文档写入器
- `CreateNewDocument() error` - 创建新文档
- `AddParagraph(text, style string) error` - 添加段落
- `AddTable(data [][]string) error` - 添加表格
- `Save(filename string) error` - 保存文档

#### `pkg/opc` - OPC 容器处理
- `Open(filename string) (*Container, error)` - 打开 OPC 容器
- `GetPart(uri string) (*Part, error)` - 获取文档部分
- `GetParts() map[string]*Part` - 获取所有部分

#### `pkg/parser` - XML 解析器
- `ParseWordML(data []byte) (*types.DocumentContent, error)` - 解析 WordML
- `ParseXML(data []byte, v interface{}) error` - 通用 XML 解析

### 高级功能

#### 文档格式化
```go
// 使用高级格式化功能
formatter := wordprocessingml.NewAdvancedFormatter(doc)

// 创建复杂表格
table := formatter.CreateComplexTable(3, 3)

// 合并单元格
err := formatter.MergeCells(table, "A1", "B2")

// 设置单元格边框
err = formatter.SetCellBorders(table, "A1", "solid", "black", 1)
```

#### 文档保护
```go
// 设置文档保护
protector := wordprocessingml.NewDocumentProtector(doc)
err := protector.SetPassword("password123")
err = protector.ProtectDocument("readOnly")
```

## 📁 项目结构

```
go-word/
├── pkg/
│   ├── opc/              # OPC 容器处理
│   │   └── package.go    # OPC 包定义
│   ├── wordprocessingml/ # Word 文档处理
│   │   ├── document.go   # 文档核心功能
│   │   ├── advanced_formatting.go  # 高级格式化
│   │   ├── document_parts.go       # 文档部分
│   │   └── document_protection.go  # 文档保护
│   ├── parser/           # XML 解析器
│   │   ├── wordml.go     # WordML 解析
│   │   └── xml.go        # 通用 XML 解析
│   ├── writer/           # 文档写入器
│   │   └── document_writer.go
│   ├── types/            # 共享类型定义
│   │   └── types.go      # 核心数据结构
│   └── utils/            # 工具函数
│       └── errors.go     # 错误处理
├── examples/             # 使用示例
│   ├── basic_usage.go    # 基本用法
│   ├── advanced_usage.go # 高级用法
│   └── ...               # 更多示例
├── docs/                # 文档
│   └── PROJECT_SUMMARY.md
├── tests/               # 测试文件
└── README.md           # 项目说明
```

## 🧪 测试

### 运行所有测试
```bash
go test ./...
```

### 运行覆盖率测试
```bash
go test -cover ./pkg/...
```

### 运行性能基准
```bash
go test -bench=. ./tests/
```

## 📊 性能特性

- **内存效率**: 流式解析，最小化内存占用
- **解析速度**: 优化的 XML 解析算法
- **并发安全**: 支持并发文档处理
- **错误恢复**: 优雅的错误处理和资源清理

## 🔧 开发状态

### ✅ 已完成功能
- [x] 项目初始化和基础架构
- [x] OPC 容器基础功能
- [x] WordprocessingML 解析
- [x] 文档内容提取
- [x] 样式和格式解析
- [x] 文档修改功能
- [x] 文档创建功能
- [x] 格式化和样式修改
- [x] 高级表格操作
- [x] 文档保护功能
- [x] 完整的测试覆盖

### 🚧 开发中功能
- [ ] 模板处理系统
- [ ] 批量文档操作
- [ ] 更多格式支持
- [ ] 性能优化

### 📋 计划功能
- [ ] 图表支持
- [ ] 图片处理
- [ ] 宏支持
- [ ] 插件系统

## 🤝 贡献

我们欢迎所有形式的贡献！

### 贡献指南

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

### 开发环境设置

```bash
# 克隆项目
git clone https://github.com/tanqiangyes/go-word.git
cd go-word

# 安装依赖
go mod tidy

# 运行测试
go test ./...

# 运行示例
go run examples/basic_usage.go
```

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 📞 支持

- 📧 邮箱: [your-email@example.com]
- 🐛 问题报告: [GitHub Issues](https://github.com/tanqiangyes/go-word/issues)
- 📖 文档: [项目文档](docs/)
- 💬 讨论: [GitHub Discussions](https://github.com/tanqiangyes/go-word/discussions)

## 🙏 致谢

- Microsoft Open XML SDK 团队
- Go 语言社区
- 所有贡献者和用户

---

**⭐ 如果这个项目对您有帮助，请给我们一个星标！** 