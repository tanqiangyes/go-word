# 快速开始指南

## 概述

本指南将帮助您快速上手 Go OpenXML SDK，在几分钟内开始处理 Word 文档。

## 目录

- [安装](#安装)
- [第一个示例](#第一个示例)
- [基本操作](#基本操作)
- [高级功能](#高级功能)
- [常见问题](#常见问题)

## 安装

### 1. 安装 Go

确保您已安装 Go 1.18 或更高版本：

```bash
go version
```

如果未安装，请访问 [golang.org/dl](https://golang.org/dl) 下载安装。

### 2. 安装库

```bash
go get github.com/tanqiangyes/go-word
```

### 3. 验证安装

创建一个测试文件 `test.go`：

```go
package main

import (
    "fmt"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    fmt.Println("Go OpenXML SDK 安装成功!")
}
```

运行测试：

```bash
go run test.go
```

## 第一个示例

### 读取 Word 文档

创建一个 `first_example.go` 文件：

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // 打开 Word 文档
    doc, err := word.Open("your_document.docx")
    if err != nil {
        log.Fatal("无法打开文档:", err)
    }
    defer doc.Close()

    // 获取文档文本
    text, err := doc.GetText()
    if err != nil {
        log.Fatal("无法获取文本:", err)
    }

    fmt.Println("文档内容:")
    fmt.Println(text)
}
```

运行示例：

```bash
go run first_example.go
```

## 基本操作

### 1. 打开和关闭文档

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // 打开文档
    doc, err := word.Open("document.docx")
    if err != nil {
        log.Fatal("无法打开文档:", err)
    }
    defer doc.Close() // 确保资源被释放

    fmt.Println("文档打开成功!")
}
```

### 2. 读取文档内容

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    doc, err := word.Open("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // 获取纯文本
    text, err := doc.GetText()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Println("文档文本:", text)

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
        for j, row := range table.Rows {
            for k, cell := range row.Cells {
                fmt.Printf("  单元格[%d,%d]: %s\n", j, k, cell.Text)
            }
        }
    }
}
```

### 3. 创建新文档

```go
package main

import (
    "log"
    "github.com/tanqiangyes/go-word/pkg/writer"
    "github.com/tanqiangyes/go-word/pkg/types"
)

func main() {
    // 创建文档写入器
    docWriter := writer.NewDocumentWriter()

    // 创建新文档
    err := docWriter.CreateNewDocument()
    if err != nil {
        log.Fatal(err)
    }

    // 添加段落
    err = docWriter.AddParagraph("这是一个段落", "Normal")
    if err != nil {
        log.Fatal(err)
    }

    // 添加格式化段落
    formattedRuns := []types.Run{
        {Text: "粗体文本", Bold: true, FontSize: 16},
        {Text: "斜体文本", Italic: true, FontSize: 14},
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
    err = docWriter.Save("output.docx")
    if err != nil {
        log.Fatal(err)
    }

    fmt.Println("文档创建成功!")
}
```
// 打开文档
doc, err := word.Open("document.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close() // 重要：确保资源被释放
```

### 2. 提取文本内容

```go
// 获取纯文本
text, err := doc.GetText()
if err != nil {
    log.Fatal(err)
}
fmt.Println("文档文本:", text)
```

### 3. 获取段落信息

```go
// 获取所有段落
paragraphs, err := doc.GetParagraphs()
if err != nil {
    log.Fatal(err)
}

for i, paragraph := range paragraphs {
    fmt.Printf("段落 %d: %s\n", i+1, paragraph.Text)
    
    // 显示段落中的格式化信息
    for j, run := range paragraph.Runs {
        fmt.Printf("  运行 %d: '%s'", j+1, run.Text)
        if run.Bold {
            fmt.Print(" [粗体]")
        }
        if run.Italic {
            fmt.Print(" [斜体]")
        }
        fmt.Println()
    }
}
```

### 4. 获取表格信息

```go
// 获取所有表格
tables, err := doc.GetTables()
if err != nil {
    log.Fatal(err)
}

for i, table := range tables {
    fmt.Printf("表格 %d: %d行 x %d列\n", i+1, len(table.Rows), table.Columns)
    
    // 显示表格内容
    for rowIdx, row := range table.Rows {
        fmt.Printf("  行 %d: ", rowIdx+1)
        for colIdx, cell := range row.Cells {
            if colIdx > 0 {
                fmt.Print(" | ")
            }
            fmt.Print(cell.Text)
        }
        fmt.Println()
    }
}
```

### 5. 创建新文档

```go
package main

import (
    "log"
    "github.com/tanqiangyes/go-word/pkg/writer"
    "github.com/tanqiangyes/go-word/pkg/types"
)

func main() {
    // 创建文档写入器
    docWriter := writer.NewDocumentWriter()
    
    // 创建新文档
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
            Text:     "普通文本",
            FontSize: 12,
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
    
    fmt.Println("文档创建成功: new_document.docx")
}
```

## 高级功能

### 1. 高级表格操作

```go
// 创建高级格式化器
formatter := word.NewAdvancedFormatter(doc)

// 创建复杂表格
table := formatter.CreateComplexTable(3, 3)

// 合并单元格
err := formatter.MergeCells(table, "A1", "B2")
if err != nil {
    log.Fatal(err)
}

// 设置单元格边框
err = formatter.SetCellBorders(table, "A1", "solid", "black", 1)
if err != nil {
    log.Fatal(err)
}

// 设置单元格背景色
err = formatter.SetCellShading(table, "A1", "#FF0000")
if err != nil {
    log.Fatal(err)
}
```

### 2. 文档保护

```go
// 创建文档保护器
protector := word.NewDocumentProtector(doc)

// 设置密码
err := protector.SetPassword("password123")
if err != nil {
    log.Fatal(err)
}

// 保护文档
err = protector.ProtectDocument("readOnly")
if err != nil {
    log.Fatal(err)
}
```

### 3. 批量处理

```go
// 批量处理多个文档
filenames := []string{"doc1.docx", "doc2.docx", "doc3.docx"}

for _, filename := range filenames {
    doc, err := word.Open(filename)
    if err != nil {
        log.Printf("无法打开 %s: %v", filename, err)
        continue
    }
    
    // 处理文档
    text, err := doc.GetText()
    if err != nil {
        log.Printf("无法获取 %s 的文本: %v", filename, err)
        doc.Close()
        continue
    }
    
    fmt.Printf("文档 %s: %d 字符\n", filename, len(text))
    
    // 关闭文档
    doc.Close()
}
```

## 常见问题

### Q: 如何检查文档是否为空？

```go
text, err := doc.GetText()
if err != nil {
    log.Fatal(err)
}

if text == "" {
    fmt.Println("文档为空")
} else {
    fmt.Printf("文档包含 %d 个字符\n", len(text))
}
```

### Q: 如何处理大文档？

```go
// 对于大文档，避免一次性加载所有内容
paragraphs, err := doc.GetParagraphs()
if err != nil {
    log.Fatal(err)
}

// 分批处理段落
for i, paragraph := range paragraphs {
    if i%100 == 0 {
        fmt.Printf("处理进度: %d/%d\n", i, len(paragraphs))
    }
    
    // 处理段落...
    processParagraph(paragraph)
}
```

### Q: 如何检查文档格式？

```go
// 检查文档是否包含表格
tables, err := doc.GetTables()
if err != nil {
    log.Fatal(err)
}

if len(tables) > 0 {
    fmt.Printf("文档包含 %d 个表格\n", len(tables))
} else {
    fmt.Println("文档不包含表格")
}

// 检查文档是否包含格式化文本
paragraphs, err := doc.GetParagraphs()
if err != nil {
    log.Fatal(err)
}

hasFormatting := false
for _, paragraph := range paragraphs {
    for _, run := range paragraph.Runs {
        if run.Bold || run.Italic || run.Underline {
            hasFormatting = true
            break
        }
    }
    if hasFormatting {
        break
    }
}

if hasFormatting {
    fmt.Println("文档包含格式化文本")
} else {
    fmt.Println("文档不包含格式化文本")
}
```

### Q: 如何处理错误？

```go
doc, err := word.Open("document.docx")
if err != nil {
    // 检查错误类型
    switch {
    case errors.Is(err, &DocumentError{}):
        fmt.Println("文档格式错误")
    case errors.Is(err, &ParseError{}):
        fmt.Println("解析错误")
    default:
        fmt.Printf("未知错误: %v\n", err)
    }
    return
}
defer doc.Close()
```

### Q: 如何获取文档统计信息？

```go
// 获取文档统计信息
func getDocumentStats(doc *word.Document) {
    paragraphs, _ := doc.GetParagraphs()
    tables, _ := doc.GetTables()
    text, _ := doc.GetText()
    
    fmt.Printf("文档统计:\n")
    fmt.Printf("  - 段落数: %d\n", len(paragraphs))
    fmt.Printf("  - 表格数: %d\n", len(tables))
    fmt.Printf("  - 字符数: %d\n", len(text))
    
    // 计算单词数
    words := strings.Fields(text)
    fmt.Printf("  - 单词数: %d\n", len(words))
    
    // 计算行数
    lines := strings.Split(text, "\n")
    fmt.Printf("  - 行数: %d\n", len(lines))
}
```

## 高级功能示例

### 1. 文档质量改进

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // 打开文档
    doc, err := word.Open("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // 创建质量改进管理器
    manager := word.NewDocumentQualityManager(doc)

    // 改进文档质量
    err = manager.ImproveDocumentQuality()
    if err != nil {
        log.Fatal(err)
    }

    // 获取质量报告
    report := manager.GetQualityReport()
    fmt.Println("质量报告:")
    fmt.Println(report)
}
```

### 2. 高级样式系统

```go
package main

import (
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
    "github.com/tanqiangyes/go-word/pkg/types"
)

func main() {
    // 创建样式系统
    system := word.NewAdvancedStyleSystem()

    // 定义标题样式
    headingStyle := &word.ParagraphStyleDefinition{
        ID:   "Heading1",
        Name: "Heading 1",
        BasedOn: "Normal",
        Properties: &word.ParagraphStyleProperties{
            Alignment: "left",
        },
    }

    // 添加样式
    err := system.AddParagraphStyle(headingStyle)
    if err != nil {
        log.Fatal(err)
    }

    // 应用样式到段落
    paragraph := &types.Paragraph{Text: "这是一个标题"}
    err = system.ApplyStyle(paragraph, "Heading1")
    if err != nil {
        log.Fatal(err)
    }

    fmt.Println("样式应用成功!")
}
```

### 3. 文档保护

```go
package main

import (
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // 创建文档保护
    protection := word.NewDocumentProtection()

    // 启用只读保护
    err := protection.EnableProtection(word.ReadOnlyProtection, "password123")
    if err != nil {
        log.Fatal(err)
    }

    fmt.Println("文档保护已启用!")

    // 禁用保护
    err = protection.DisableProtection()
    if err != nil {
        log.Fatal(err)
    }

    fmt.Println("文档保护已禁用!")
}
```

### 4. 文档验证

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // 打开文档
    doc, err := word.Open("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // 创建验证器
    validator := word.NewDocumentValidator()

    // 添加验证规则
    rule := word.ValidationRule{
        ID: "check_spelling",
        Name: "拼写检查",
        Type: word.SpellingRule,
        Enabled: true,
    }

    err = validator.AddRule(rule)
    if err != nil {
        log.Fatal(err)
    }

    // 验证文档
    result, err := validator.ValidateDocument(doc)
    if err != nil {
        log.Fatal(err)
    }

    if result.IsValid {
        fmt.Println("文档验证通过!")
    } else {
        fmt.Printf("发现 %d 个问题\n", len(result.Issues))
    }
}
```

### 5. 批处理

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // 创建批处理器
    processor := word.NewBatchProcessor(4)

    // 添加处理任务
    filenames := []string{"doc1.docx", "doc2.docx", "doc3.docx"}
    
    for _, filename := range filenames {
        processor.AddTask(filename, func(doc *word.Document) error {
            // 处理文档
            text, err := doc.GetText()
            if err != nil {
                return err
            }
            
            fmt.Printf("处理文档 %s: %d 字符\n", filename, len(text))
            return nil
        })
    }

    // 开始处理
    err := processor.Process()
    if err != nil {
        log.Fatal(err)
    }

    fmt.Println("批处理完成!")
}
```

### 6. 错误处理最佳实践

```go
package main

import (
    "fmt"
    "log"
    "errors"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // 使用结构化错误处理
    doc, err := word.Open("document.docx")
    if err != nil {
        var docErr *word.DocumentError
        if errors.As(err, &docErr) {
            switch docErr.Code {
            case "FILE_NOT_FOUND":
                fmt.Println("文件未找到:", docErr.Message)
            case "INVALID_FORMAT":
                fmt.Println("文件格式无效:", docErr.Message)
            default:
                fmt.Println("未知错误:", docErr.Message)
            }
        }
        log.Fatal(err)
    }
    defer doc.Close()

    // 处理文档...
}
```

### 7. 性能优化示例

```go
package main

import (
    "fmt"
    "log"
    "runtime"
    "sync"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // 并发处理多个文档
    filenames := []string{"doc1.docx", "doc2.docx", "doc3.docx", "doc4.docx"}
    
    var wg sync.WaitGroup
    semaphore := make(chan struct{}, runtime.NumCPU())
    
    for _, filename := range filenames {
        wg.Add(1)
        go func(fname string) {
            defer wg.Done()
            semaphore <- struct{}{} // 获取信号量
            defer func() { <-semaphore }() // 释放信号量
            
            doc, err := word.Open(fname)
            if err != nil {
                log.Printf("处理文件 %s 时出错: %v", fname, err)
                return
            }
            defer doc.Close()
            
            // 处理文档
            text, err := doc.GetText()
            if err != nil {
                log.Printf("获取文本失败: %v", err)
                return
            }
            
            fmt.Printf("文档 %s: %d 字符\n", fname, len(text))
        }(filename)
    }
    
    wg.Wait()
    fmt.Println("所有文档处理完成!")
}
```

## 下一步

现在您已经掌握了基本用法，可以：

1. **查看完整示例**: 运行 `go run examples/basic_usage.go`
2. **阅读 API 文档**: 查看 `docs/API_REFERENCE.md`
3. **学习高级功能**: 查看 `examples/advanced_usage.go`
4. **参与开发**: 查看 `docs/DEVELOPMENT_GUIDE.md`

## 获取帮助

- 📖 **文档**: [API 参考](docs/API_REFERENCE.md)
- 🐛 **问题报告**: [GitHub Issues](https://github.com/tanqiangyes/go-word/issues)
- 💬 **讨论**: [GitHub Discussions](https://github.com/tanqiangyes/go-word/discussions)
- 📧 **邮箱**: [your-email@example.com]

---

**提示**: 如果您在使用过程中遇到问题，请先查看常见问题部分，如果问题仍未解决，请创建 GitHub Issue。 