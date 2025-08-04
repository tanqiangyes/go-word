# Go OpenXML SDK

一个用 Go 语言重写的 Microsoft Open XML SDK，专门用于 Word 文档处理。

## 特性

- 🚀 **高性能**：优化的解析速度，最小化内存占用
- 📄 **完整支持**：支持 Word 文档格式（.docx），包含新老版本
- 🔧 **Go 风格**：遵循 Go 语言最佳实践和惯用法
- 📚 **详细文档**：完整的 API 文档和使用示例
- 🛡️ **错误处理**：详细的错误信息和日志记录

## 快速开始

### 安装

```bash
go get github.com/go-word
```

### 基本使用

#### 读取文档

```go
package main

import (
    "fmt"
    "log"

    "github.com/go-word/pkg/wordprocessingml"
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

    "github.com/go-word/pkg/writer"
    "github.com/go-word/pkg/types"
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

## 项目结构

```
go-word/
├── pkg/
│   ├── opc/              # OPC 容器处理
│   ├── wordprocessingml/ # Word 文档处理
│   ├── parser/           # 解析器
│   ├── writer/           # 写入器
│   └── utils/            # 工具函数
├── examples/             # 使用示例
├── docs/                # 文档
└── tests/               # 测试
```

## 开发状态

- [x] 项目初始化
- [x] OPC 容器基础功能
- [x] WordprocessingML 解析
- [x] 文档内容提取
- [x] 样式和格式解析
- [x] 文档修改功能
- [x] 文档创建功能
- [x] 格式化和样式修改
- [ ] 高级操作功能
- [ ] 模板处理
- [ ] 批量操作

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可证

MIT License 