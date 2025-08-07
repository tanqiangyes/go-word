# Go Word API 文档

## 概述

Go Word 是一个用 Go 语言编写的 Microsoft Open XML SDK，专门用于 Word 文档处理。本文档提供了完整的 API 参考和使用指南。

## 包结构

```
pkg/
├── wordprocessingml/    # 核心文档处理
├── writer/             # 文档写入器
├── parser/             # XML 解析器
├── types/              # 共享类型定义
├── utils/              # 工具函数
└── opc/                # Open Packaging Conventions
```

## 核心 API

### Document 结构

`Document` 是文档的核心结构，代表一个打开的 Word 文档。

```go
type Document struct {
    mainPart *MainDocumentPart
    parts    *DocumentParts
}
```

#### 方法

##### Open(filename string) (*Document, error)
打开一个 Word 文档文件。

```go
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close()
```

##### Close() error
关闭文档并释放资源。

##### GetText() (string, error)
获取文档的纯文本内容。

```go
text, err := doc.GetText()
if err != nil {
    log.Fatal(err)
}
fmt.Println(text)
```

##### GetParagraphs() ([]types.Paragraph, error)
获取文档中的所有段落。

```go
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
```

##### GetTables() ([]types.Table, error)
获取文档中的所有表格。

```go
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
```

### DocumentWriter

`DocumentWriter` 用于创建和修改 Word 文档。

```go
type DocumentWriter struct {
    document *Document
}
```

#### 方法

##### NewDocumentWriter() *DocumentWriter
创建新的文档写入器。

```go
writer := writer.NewDocumentWriter()
```

##### CreateNewDocument() error
创建新的空白文档。

```go
err := writer.CreateNewDocument()
if err != nil {
    log.Fatal(err)
}
```

##### AddParagraph(text, style string) error
添加段落到文档。

```go
err := writer.AddParagraph("这是一个段落", "Normal")
if err != nil {
    log.Fatal(err)
}
```

##### AddFormattedParagraph(text, style string, runs []types.Run) error
添加带格式的段落。

```go
formattedRuns := []types.Run{
    {Text: "粗体文本", Bold: true, FontSize: 16},
    {Text: "斜体文本", Italic: true, FontSize: 14},
}

err := writer.AddFormattedParagraph("格式化段落", "Normal", formattedRuns)
if err != nil {
    log.Fatal(err)
}
```

##### AddTable(data [][]string) error
添加表格到文档。

```go
tableData := [][]string{
    {"姓名", "年龄", "职业"},
    {"张三", "25", "工程师"},
    {"李四", "30", "设计师"},
}

err := writer.AddTable(tableData)
if err != nil {
    log.Fatal(err)
}
```

##### Save(filename string) error
保存文档到文件。

```go
err := writer.Save("output.docx")
if err != nil {
    log.Fatal(err)
}
```

### 类型定义

#### Paragraph

```go
type Paragraph struct {
    Text       string
    Style      string
    Runs       []Run
    HasComment bool
    CommentID  string
}
```

#### Run

```go
type Run struct {
    Text      string
    Bold      bool
    Italic    bool
    Underline bool
    FontSize  int
    FontName  string
    Color     string
}
```

#### Table

```go
type Table struct {
    Rows    []TableRow
    Columns int
}
```

#### TableRow

```go
type TableRow struct {
    Cells []TableCell
}
```

#### TableCell

```go
type TableCell struct {
    Text string
}
```

## 高级功能

### 文档质量管理系统

#### DocumentQualityManager

```go
type DocumentQualityManager struct {
    Document *Document
    Settings *QualitySettings
    Metrics  *QualityMetrics
}
```

##### NewDocumentQualityManager(doc *Document) *DocumentQualityManager
创建文档质量管理器。

```go
manager := wordprocessingml.NewDocumentQualityManager(doc)
```

##### ImproveDocumentQuality() error
改进文档质量。

```go
err := manager.ImproveDocumentQuality()
if err != nil {
    log.Fatal(err)
}
```

##### GetQualityReport() string
获取质量报告。

```go
report := manager.GetQualityReport()
fmt.Println(report)
```

### 高级样式系统

#### AdvancedStyleSystem

```go
type AdvancedStyleSystem struct {
    StyleManager      *StyleManager
    StyleCache        map[string]*StyleDefinition
    InheritanceChain  map[string][]string
    ConflictResolver  *StyleConflictResolver
}
```

##### NewAdvancedStyleSystem() *AdvancedStyleSystem
创建高级样式系统。

```go
system := wordprocessingml.NewAdvancedStyleSystem()
```

##### AddParagraphStyle(style *ParagraphStyleDefinition) error
添加段落样式。

```go
style := &wordprocessingml.ParagraphStyleDefinition{
    ID:   "Heading1",
    Name: "Heading 1",
    BasedOn: "Normal",
    Properties: &wordprocessingml.ParagraphStyleProperties{
        Alignment: "left",
    },
}

err := system.AddParagraphStyle(style)
if err != nil {
    log.Fatal(err)
}
```

##### ApplyStyle(content interface{}, styleID string) error
应用样式到内容。

```go
paragraph := &types.Paragraph{Text: "测试段落"}
err := system.ApplyStyle(paragraph, "Heading1")
if err != nil {
    log.Fatal(err)
}
```

### 文档保护

#### DocumentProtection

```go
type DocumentProtection struct {
    Enabled     bool
    Password    string
    ProtectionType ProtectionType
    Enforcement  EnforcementLevel
}
```

##### EnableProtection(protectionType ProtectionType, password string) error
启用文档保护。

```go
protection := wordprocessingml.NewDocumentProtection()
err := protection.EnableProtection(wordprocessingml.ReadOnlyProtection, "password123")
if err != nil {
    log.Fatal(err)
}
```

##### DisableProtection() error
禁用文档保护。

```go
err := protection.DisableProtection()
if err != nil {
    log.Fatal(err)
}
```

### 文档验证

#### DocumentValidator

```go
type DocumentValidator struct {
    Rules    []ValidationRule
    AutoFix  bool
    Settings *ValidationSettings
}
```

##### NewDocumentValidator() *DocumentValidator
创建文档验证器。

```go
validator := wordprocessingml.NewDocumentValidator()
```

##### AddRule(rule ValidationRule) error
添加验证规则。

```go
rule := wordprocessingml.ValidationRule{
    ID: "check_spelling",
    Name: "拼写检查",
    Type: wordprocessingml.SpellingRule,
    Enabled: true,
}

err := validator.AddRule(rule)
if err != nil {
    log.Fatal(err)
}
```

##### ValidateDocument(doc *Document) (*ValidationResult, error)
验证文档。

```go
result, err := validator.ValidateDocument(doc)
if err != nil {
    log.Fatal(err)
}

if result.IsValid {
    fmt.Println("文档验证通过")
} else {
    fmt.Printf("发现 %d 个问题\n", len(result.Issues))
}
```

## 错误处理

### 错误类型

```go
type DocumentError struct {
    Code     string
    Message  string
    Severity ErrorSeverity
    Context  map[string]interface{}
}
```

### 错误处理示例

```go
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    var docErr *wordprocessingml.DocumentError
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
```

## 性能优化

### 内存管理

```go
// 使用 defer 确保资源释放
doc, err := wordprocessingml.Open("large_document.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close()

// 分批处理大文档
const batchSize = 1000
paragraphs, err := doc.GetParagraphs()
if err != nil {
    log.Fatal(err)
}

for i := 0; i < len(paragraphs); i += batchSize {
    end := i + batchSize
    if end > len(paragraphs) {
        end = len(paragraphs)
    }
    
    batch := paragraphs[i:end]
    // 处理批次
    processBatch(batch)
}
```

### 并发处理

```go
// 并发处理多个文档
func processDocuments(filenames []string) {
    semaphore := make(chan struct{}, 4) // 限制并发数
    var wg sync.WaitGroup
    
    for _, filename := range filenames {
        wg.Add(1)
        go func(fname string) {
            defer wg.Done()
            semaphore <- struct{}{} // 获取信号量
            defer func() { <-semaphore }() // 释放信号量
            
            doc, err := wordprocessingml.Open(fname)
            if err != nil {
                log.Printf("处理文件 %s 时出错: %v", fname, err)
                return
            }
            defer doc.Close()
            
            // 处理文档
            processDocument(doc)
        }(filename)
    }
    
    wg.Wait()
}
```

## 最佳实践

### 1. 资源管理

```go
// 正确的方式
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close() // 确保资源释放

// 错误的方式
doc, _ := wordprocessingml.Open("document.docx")
// 没有调用 Close()
```

### 2. 错误处理

```go
// 正确的方式
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    log.Fatal(err)
}

// 错误的方式
doc, _ := wordprocessingml.Open("document.docx")
// 忽略错误
```

### 3. 性能考虑

```go
// 对于大文档，使用流式处理
doc, err := wordprocessingml.Open("large_document.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close()

// 分批处理段落
paragraphs, err := doc.GetParagraphs()
if err != nil {
    log.Fatal(err)
}

const batchSize = 1000
for i := 0; i < len(paragraphs); i += batchSize {
    end := i + batchSize
    if end > len(paragraphs) {
        end = len(paragraphs)
    }
    
    batch := paragraphs[i:end]
    processBatch(batch)
}
```

### 4. 样式使用

```go
// 创建样式系统
system := wordprocessingml.NewAdvancedStyleSystem()

// 定义样式
headingStyle := &wordprocessingml.ParagraphStyleDefinition{
    ID:   "Heading1",
    Name: "Heading 1",
    BasedOn: "Normal",
    Properties: &wordprocessingml.ParagraphStyleProperties{
        Alignment: "left",
    },
}

err := system.AddParagraphStyle(headingStyle)
if err != nil {
    log.Fatal(err)
}

// 应用样式
paragraph := &types.Paragraph{Text: "标题"}
err = system.ApplyStyle(paragraph, "Heading1")
if err != nil {
    log.Fatal(err)
}
```

## 常见问题

### Q: 如何处理大文档？

A: 使用分批处理和内存管理技术：

```go
// 分批读取段落
const batchSize = 1000
for i := 0; i < len(paragraphs); i += batchSize {
    end := i + batchSize
    if end > len(paragraphs) {
        end = len(paragraphs)
    }
    
    batch := paragraphs[i:end]
    processBatch(batch)
}
```

### Q: 如何提高性能？

A: 使用并发处理和缓存：

```go
// 并发处理多个文档
func processDocuments(filenames []string) {
    var wg sync.WaitGroup
    semaphore := make(chan struct{}, runtime.NumCPU())
    
    for _, filename := range filenames {
        wg.Add(1)
        go func(fname string) {
            defer wg.Done()
            semaphore <- struct{}{}
            defer func() { <-semaphore }()
            
            processDocument(fname)
        }(filename)
    }
    
    wg.Wait()
}
```

### Q: 如何处理特殊字符？

A: 使用编码处理：

```go
// 处理特殊字符
text := "包含特殊字符的文本：< > & \" '"
// 库会自动处理 XML 特殊字符
err := writer.AddParagraph(text, "Normal")
if err != nil {
    log.Fatal(err)
}
```

## 版本兼容性

### 支持的 Word 版本

- Microsoft Word 2007 (.docx)
- Microsoft Word 2010 (.docx)
- Microsoft Word 2013 (.docx)
- Microsoft Word 2016 (.docx)
- Microsoft Word 2019 (.docx)
- Microsoft Word 365 (.docx)

### 兼容性注意事项

1. **向后兼容性**: 新版本生成的文档可以在旧版本 Word 中打开
2. **功能兼容性**: 某些高级功能在旧版本中可能不可用
3. **格式兼容性**: 建议在目标 Word 版本中测试文档

## 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。 