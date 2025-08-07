# API 参考文档

## 概述

本文档提供了 Go OpenXML SDK 的完整 API 参考。该库提供了高性能的 Word 文档处理功能，包括读取、解析、修改和创建 Word 文档。

## 目录

- [核心包](#核心包)
  - [wordprocessingml](#wordprocessingml)
  - [writer](#writer)
  - [opc](#opc)
  - [parser](#parser)
  - [types](#types)
  - [utils](#utils)
- [高级功能](#高级功能)
- [错误处理](#错误处理)
- [最佳实践](#最佳实践)

## 核心包

### wordprocessingml

Word 文档处理的核心包，提供文档的读取、解析和操作功能。

#### 类型

##### Document
```go
type Document struct {
    container *opc.Container
    mainPart  *MainDocumentPart
    parts     map[string]*opc.Part
    documentParts *DocumentParts
}
```

表示一个 Word 文档，包含文档的所有内容和元数据。

##### MainDocumentPart
```go
type MainDocumentPart struct {
    Content *types.DocumentContent
    DocumentProperties map[string]interface{}
    Relationships []Relationship
}
```

表示文档的主要部分，包含文档的实际内容。

##### DocumentQualityManager
```go
type DocumentQualityManager struct {
    Document *Document
    Settings *QualitySettings
    Metrics  *QualityMetrics
}
```

文档质量管理系统，用于改进文档质量。

##### AdvancedStyleSystem
```go
type AdvancedStyleSystem struct {
    StyleManager      *StyleManager
    StyleCache        map[string]*StyleDefinition
    InheritanceChain  map[string][]string
    ConflictResolver  *StyleConflictResolver
}
```

高级样式系统，管理文档样式和格式。

#### 函数

##### Open
```go
func Open(filename string) (*Document, error)
```

打开一个 Word 文档文件。

**参数:**
- `filename string`: 文档文件路径

**返回值:**
- `*Document`: 文档对象
- `error`: 错误信息

**示例:**
```go
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close()
```

##### Close
```go
func (d *Document) Close() error
```

关闭文档并释放资源。

**返回值:**
- `error`: 错误信息

##### GetText
```go
func (d *Document) GetText() (string, error)
```

获取文档的纯文本内容。

**返回值:**
- `string`: 文档的纯文本内容
- `error`: 错误信息

**示例:**
```go
text, err := doc.GetText()
if err != nil {
    log.Fatal(err)
}
fmt.Println("文档内容:", text)
```

##### GetParagraphs
```go
func (d *Document) GetParagraphs() ([]Paragraph, error)
```

获取文档中的所有段落。

**返回值:**
- `[]Paragraph`: 段落列表
- `error`: 错误信息

**示例:**
```go
paragraphs, err := doc.GetParagraphs()
if err != nil {
    log.Fatal(err)
}

for i, paragraph := range paragraphs {
    fmt.Printf("段落 %d: %s\n", i+1, paragraph.Text)
}
```

##### GetTables
```go
func (d *Document) GetTables() ([]Table, error)
```

获取文档中的所有表格。

**返回值:**
- `[]Table`: 表格列表
- `error`: 错误信息

**示例:**
```go
tables, err := doc.GetTables()
if err != nil {
    log.Fatal(err)
}

for i, table := range tables {
    fmt.Printf("表格 %d: %d行 x %d列\n", i+1, len(table.Rows), table.Columns)
}
```

##### GetContainer
```go
func (d *Document) GetContainer() *opc.Container
```

获取底层的 OPC 容器。

**返回值:**
- `*opc.Container`: OPC 容器对象

##### GetDocumentParts
```go
func (d *Document) GetDocumentParts() *DocumentParts
```

获取文档部分管理器。

**返回值:**
- `*DocumentParts`: 文档部分对象

#### 高级格式化

##### NewAdvancedFormatter
```go
func NewAdvancedFormatter(doc *Document) *AdvancedFormatter
```

创建高级格式化器。

**参数:**
- `doc *Document`: 文档对象

**返回值:**
- `*AdvancedFormatter`: 高级格式化器

##### CreateComplexTable
```go
func (af *AdvancedFormatter) CreateComplexTable(rows, cols int) *Table
```

创建复杂表格。

**参数:**
- `rows int`: 行数
- `cols int`: 列数

**返回值:**
- `*Table`: 创建的表格

**示例:**
```go
formatter := wordprocessingml.NewAdvancedFormatter(doc)
table := formatter.CreateComplexTable(3, 3)
```

##### MergeCells
```go
func (af *AdvancedFormatter) MergeCells(table *Table, startCell, endCell string) error
```

合并表格单元格。

**参数:**
- `table *Table`: 表格对象
- `startCell string`: 起始单元格引用（如 "A1"）
- `endCell string`: 结束单元格引用（如 "B2"）

**返回值:**
- `error`: 错误信息

**示例:**
```go
err := formatter.MergeCells(table, "A1", "B2")
if err != nil {
    log.Fatal(err)
}
```

##### SetCellBorders
```go
func (af *AdvancedFormatter) SetCellBorders(table *Table, cellRef, style, color string, width int) error
```

设置单元格边框。

**参数:**
- `table *Table`: 表格对象
- `cellRef string`: 单元格引用
- `style string`: 边框样式
- `color string`: 边框颜色
- `width int`: 边框宽度

**返回值:**
- `error`: 错误信息

**示例:**
```go
err := formatter.SetCellBorders(table, "A1", "solid", "black", 1)
if err != nil {
    log.Fatal(err)
}
```

##### SetCellShading
```go
func (af *AdvancedFormatter) SetCellShading(table *Table, cellRef, color string) error
```

设置单元格背景色。

**参数:**
- `table *Table`: 表格对象
- `cellRef string`: 单元格引用
- `color string`: 背景颜色

**返回值:**
- `error`: 错误信息

**示例:**
```go
err := formatter.SetCellShading(table, "A1", "#FF0000")
if err != nil {
    log.Fatal(err)
}
```

#### 文档保护

##### NewDocumentProtector
```go
func NewDocumentProtector(doc *Document) *DocumentProtector
```

创建文档保护器。

**参数:**
- `doc *Document`: 文档对象

**返回值:**
- `*DocumentProtector`: 文档保护器

##### SetPassword
```go
func (dp *DocumentProtector) SetPassword(password string) error
```

设置文档密码。

**参数:**
- `password string`: 密码

**返回值:**
- `error`: 错误信息

**示例:**
```go
protector := wordprocessingml.NewDocumentProtector(doc)
err := protector.SetPassword("password123")
if err != nil {
    log.Fatal(err)
}
```

##### ProtectDocument
```go
func (dp *DocumentProtector) ProtectDocument(protectionType string) error
```

保护文档。

**参数:**
- `protectionType string`: 保护类型（如 "readOnly"）

**返回值:**
- `error`: 错误信息

**示例:**
```go
err := protector.ProtectDocument("readOnly")
if err != nil {
    log.Fatal(err)
}
```

### writer

文档写入器包，提供文档创建和修改功能。

#### 类型

##### DocumentWriter
```go
type DocumentWriter struct {
    // 内部字段...
}
```

文档写入器，用于创建和修改 Word 文档。

#### 函数

##### NewDocumentWriter
```go
func NewDocumentWriter() *DocumentWriter
```

创建新的文档写入器。

**返回值:**
- `*DocumentWriter`: 文档写入器

**示例:**
```go
docWriter := writer.NewDocumentWriter()
```

##### CreateNewDocument
```go
func (dw *DocumentWriter) CreateNewDocument() error
```

创建新文档。

**返回值:**
- `error`: 错误信息

**示例:**
```go
err := docWriter.CreateNewDocument()
if err != nil {
    log.Fatal(err)
}
```

##### AddParagraph
```go
func (dw *DocumentWriter) AddParagraph(text, style string) error
```

添加段落。

**参数:**
- `text string`: 段落文本
- `style string`: 段落样式

**返回值:**
- `error`: 错误信息

**示例:**
```go
err := docWriter.AddParagraph("这是一个段落", "Normal")
if err != nil {
    log.Fatal(err)
}
```

##### AddFormattedParagraph
```go
func (dw *DocumentWriter) AddFormattedParagraph(text, style string, runs []types.Run) error
```

添加带格式的段落。

**参数:**
- `text string`: 段落文本
- `style string`: 段落样式
- `runs []types.Run`: 格式化的文本运行

**返回值:**
- `error`: 错误信息

**示例:**
```go
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
err := docWriter.AddFormattedParagraph("格式化段落", "Normal", formattedRuns)
if err != nil {
    log.Fatal(err)
}
```

##### AddTable
```go
func (dw *DocumentWriter) AddTable(data [][]string) error
```

添加表格。

**参数:**
- `data [][]string`: 表格数据

**返回值:**
- `error`: 错误信息

**示例:**
```go
tableData := [][]string{
    {"姓名", "年龄", "职业"},
    {"张三", "25", "工程师"},
    {"李四", "30", "设计师"},
}
err := docWriter.AddTable(tableData)
if err != nil {
    log.Fatal(err)
}
```

##### Save
```go
func (dw *DocumentWriter) Save(filename string) error
```

保存文档。

**参数:**
- `filename string`: 保存的文件名

**返回值:**
- `error`: 错误信息

**示例:**
```go
err := docWriter.Save("new_document.docx")
if err != nil {
    log.Fatal(err)
}
```

### opc

OPC（Open Packaging Convention）容器处理包。

#### 类型

##### Container
```go
type Container struct {
    // 内部字段...
}
```

OPC 容器，表示一个 Word 文档的 ZIP 容器。

##### Part
```go
type Part struct {
    URI  string
    Data []byte
    // 其他字段...
}
```

文档部分，表示容器中的一个文件。

#### 函数

##### Open
```go
func Open(filename string) (*Container, error)
```

打开 OPC 容器。

**参数:**
- `filename string`: 文件路径

**返回值:**
- `*Container`: 容器对象
- `error`: 错误信息

**示例:**
```go
container, err := opc.Open("document.docx")
if err != nil {
    log.Fatal(err)
}
defer container.Close()
```

##### Close
```go
func (c *Container) Close() error
```

关闭容器。

**返回值:**
- `error`: 错误信息

##### GetPart
```go
func (c *Container) GetPart(uri string) (*Part, error)
```

获取文档部分。

**参数:**
- `uri string`: 部分 URI

**返回值:**
- `*Part`: 部分对象
- `error`: 错误信息

**示例:**
```go
part, err := container.GetPart("word/document.xml")
if err != nil {
    log.Fatal(err)
}
```

##### GetParts
```go
func (c *Container) GetParts() map[string]*Part
```

获取所有部分。

**返回值:**
- `map[string]*Part`: 部分映射

**示例:**
```go
parts := container.GetParts()
for uri, part := range parts {
    fmt.Printf("部分: %s, 大小: %d 字节\n", uri, len(part.Data))
}
```

### parser

XML 解析器包，提供 WordML 和通用 XML 解析功能。

#### 函数

##### ParseWordML
```go
func ParseWordML(data []byte) (*types.DocumentContent, error)
```

解析 WordML 数据。

**参数:**
- `data []byte`: WordML 数据

**返回值:**
- `*types.DocumentContent`: 文档内容
- `error`: 错误信息

**示例:**
```go
content, err := parser.ParseWordML(xmlData)
if err != nil {
    log.Fatal(err)
}
```

##### ParseXML
```go
func ParseXML(data []byte, v interface{}) error
```

解析通用 XML 数据。

**参数:**
- `data []byte`: XML 数据
- `v interface{}`: 目标结构体

**返回值:**
- `error`: 错误信息

**示例:**
```go
var result MyStruct
err := parser.ParseXML(xmlData, &result)
if err != nil {
    log.Fatal(err)
}
```

### types

共享类型定义包，包含所有核心数据结构。

#### 类型

##### Paragraph
```go
type Paragraph struct {
    Text string
    Runs []Run
    Style string
    // 其他字段...
}
```

表示文档中的一个段落。

##### Run
```go
type Run struct {
    Text     string
    Bold     bool
    Italic   bool
    Underline bool
    FontSize int
    FontName string
    // 其他字段...
}
```

表示文本运行，包含格式信息。

##### Table
```go
type Table struct {
    Rows    []TableRow
    Columns int
    // 其他字段...
}
```

表示文档中的一个表格。

##### TableRow
```go
type TableRow struct {
    Cells []TableCell
    // 其他字段...
}
```

表示表格中的一行。

##### TableCell
```go
type TableCell struct {
    Text      string
    Merged    bool
    MergeStart string
    MergeEnd   string
    // 其他字段...
}
```

表示表格中的一个单元格。

##### DocumentContent
```go
type DocumentContent struct {
    Paragraphs []Paragraph
    Tables     []Table
    Text       string
    // 其他字段...
}
```

表示文档的内容结构。

### utils

工具函数包，提供错误处理和通用工具。

#### 函数

##### NewError
```go
func NewError(message string) error
```

创建新的错误。

**参数:**
- `message string`: 错误消息

**返回值:**
- `error`: 错误对象

**示例:**
```go
err := utils.NewError("文档格式无效")
```

##### WrapError
```go
func WrapError(err error, message string) error
```

包装错误。

**参数:**
- `err error`: 原始错误
- `message string`: 包装消息

**返回值:**
- `error`: 包装后的错误

**示例:**
```go
err = utils.WrapError(err, "解析文档时出错")
```

## 高级功能

### 批量处理

```go
// 批量处理多个文档
processor := wordprocessingml.NewBatchProcessor()

// 添加处理任务
processor.AddTask("document1.docx", func(doc *wordprocessingml.Document) error {
    // 处理文档
    return nil
})

// 执行批量处理
err := processor.Process()
```

### 文档验证

```go
// 验证文档结构
validator := wordprocessingml.NewDocumentValidator(doc)
err := validator.Validate()
if err != nil {
    log.Fatal(err)
}
```

### 模板处理

```go
// 使用模板创建文档
template := wordprocessingml.NewTemplate("template.docx")
doc, err := template.CreateDocument(map[string]string{
    "name": "张三",
    "age":  "25",
})
```

## 错误处理

### 错误类型

库定义了多种错误类型：

- `DocumentError`: 文档相关错误
- `ParseError`: 解析错误
- `ValidationError`: 验证错误
- `IOError`: 输入输出错误

### 错误处理最佳实践

```go
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    switch {
    case errors.Is(err, &DocumentError{}):
        log.Printf("文档错误: %v", err)
    case errors.Is(err, &ParseError{}):
        log.Printf("解析错误: %v", err)
    default:
        log.Printf("未知错误: %v", err)
    }
    return
}
defer doc.Close()
```

## 最佳实践

### 1. 资源管理

始终使用 `defer` 关闭文档：

```go
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    log.Fatal(err)
}
defer doc.Close() // 确保资源被释放
```

### 2. 错误处理

检查所有函数返回的错误：

```go
text, err := doc.GetText()
if err != nil {
    log.Printf("获取文本失败: %v", err)
    return
}
```

### 3. 性能优化

对于大文档，使用流式处理：

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
}
```

### 4. 并发安全

库支持并发使用，但需要注意：

```go
// 每个 goroutine 使用独立的文档实例
go func() {
    doc1, _ := wordprocessingml.Open("doc1.docx")
    defer doc1.Close()
    // 处理文档1
}()

go func() {
    doc2, _ := wordprocessingml.Open("doc2.docx")
    defer doc2.Close()
    // 处理文档2
}()
```

### 5. 内存管理

对于大量文档处理，注意内存使用：

```go
// 处理完一个文档后立即关闭
for _, filename := range filenames {
    doc, err := wordprocessingml.Open(filename)
    if err != nil {
        continue
    }
    
    // 处理文档
    processDocument(doc)
    
    // 立即关闭释放内存
    doc.Close()
}
```

## 版本兼容性

### Go 版本要求

- 最低版本: Go 1.18
- 推荐版本: Go 1.20+

### 平台支持

- Windows
- macOS
- Linux

### Word 文档版本支持

- Office Open XML (.docx)
- 兼容 Microsoft Word 2007 及以上版本

## 高级功能

### 文档质量管理系统

#### NewDocumentQualityManager
```go
func NewDocumentQualityManager(doc *Document) *DocumentQualityManager
```

创建文档质量管理器。

**参数:**
- `doc *Document`: 要改进的文档

**返回值:**
- `*DocumentQualityManager`: 文档质量管理器

**示例:**
```go
manager := wordprocessingml.NewDocumentQualityManager(doc)
err := manager.ImproveDocumentQuality()
if err != nil {
    log.Fatal(err)
}

report := manager.GetQualityReport()
fmt.Println(report)
```

#### ImproveDocumentQuality
```go
func (dqm *DocumentQualityManager) ImproveDocumentQuality() error
```

改进文档质量，包括元数据管理、内容质量、结构优化等。

**返回值:**
- `error`: 错误信息

#### GetQualityReport
```go
func (dqm *DocumentQualityManager) GetQualityReport() string
```

获取质量改进报告。

**返回值:**
- `string`: 质量报告

### 高级样式系统

#### NewAdvancedStyleSystem
```go
func NewAdvancedStyleSystem() *AdvancedStyleSystem
```

创建高级样式系统。

**返回值:**
- `*AdvancedStyleSystem`: 高级样式系统

**示例:**
```go
system := wordprocessingml.NewAdvancedStyleSystem()

// 添加段落样式
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

#### AddParagraphStyle
```go
func (ass *AdvancedStyleSystem) AddParagraphStyle(style *ParagraphStyleDefinition) error
```

添加段落样式。

**参数:**
- `style *ParagraphStyleDefinition`: 段落样式定义

**返回值:**
- `error`: 错误信息

#### ApplyStyle
```go
func (ass *AdvancedStyleSystem) ApplyStyle(content interface{}, styleID string) error
```

应用样式到内容。

**参数:**
- `content interface{}`: 要应用样式的内容
- `styleID string`: 样式ID

**返回值:**
- `error`: 错误信息

**示例:**
```go
paragraph := &types.Paragraph{Text: "测试段落"}
err := system.ApplyStyle(paragraph, "Heading1")
if err != nil {
    log.Fatal(err)
}
```

### 文档保护

#### NewDocumentProtection
```go
func NewDocumentProtection() *DocumentProtection
```

创建文档保护对象。

**返回值:**
- `*DocumentProtection`: 文档保护对象

#### EnableProtection
```go
func (dp *DocumentProtection) EnableProtection(protectionType ProtectionType, password string) error
```

启用文档保护。

**参数:**
- `protectionType ProtectionType`: 保护类型
- `password string`: 保护密码

**返回值:**
- `error`: 错误信息

**示例:**
```go
protection := wordprocessingml.NewDocumentProtection()
err := protection.EnableProtection(wordprocessingml.ReadOnlyProtection, "password123")
if err != nil {
    log.Fatal(err)
}
```

#### DisableProtection
```go
func (dp *DocumentProtection) DisableProtection() error
```

禁用文档保护。

**返回值:**
- `error`: 错误信息

### 文档验证

#### NewDocumentValidator
```go
func NewDocumentValidator() *DocumentValidator
```

创建文档验证器。

**返回值:**
- `*DocumentValidator`: 文档验证器

#### AddRule
```go
func (dv *DocumentValidator) AddRule(rule ValidationRule) error
```

添加验证规则。

**参数:**
- `rule ValidationRule`: 验证规则

**返回值:**
- `error`: 错误信息

**示例:**
```go
validator := wordprocessingml.NewDocumentValidator()

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

#### ValidateDocument
```go
func (dv *DocumentValidator) ValidateDocument(doc *Document) (*ValidationResult, error)
```

验证文档。

**参数:**
- `doc *Document`: 要验证的文档

**返回值:**
- `*ValidationResult`: 验证结果
- `error`: 错误信息

**示例:**
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

### 批处理功能

#### BatchProcessor
```go
type BatchProcessor struct {
    Workers int
    Timeout time.Duration
    Logger  *log.Logger
}
```

批处理器，用于处理大量文档。

#### NewBatchProcessor
```go
func NewBatchProcessor(workers int) *BatchProcessor
```

创建批处理器。

**参数:**
- `workers int`: 工作线程数

**返回值:**
- `*BatchProcessor`: 批处理器

**示例:**
```go
processor := wordprocessingml.NewBatchProcessor(4)

// 添加处理任务
processor.AddTask("doc1.docx", func(doc *Document) error {
    // 处理文档
    return nil
})

// 开始处理
err := processor.Process()
if err != nil {
    log.Fatal(err)
}
```

#### AddTask
```go
func (bp *BatchProcessor) AddTask(filename string, handler func(*Document) error)
```

添加处理任务。

**参数:**
- `filename string`: 文档文件名
- `handler func(*Document) error`: 处理函数

#### Process
```go
func (bp *BatchProcessor) Process() error
```

开始批处理。

**返回值:**
- `error`: 错误信息

---

**注意**: 本文档基于当前版本编写，API 可能会在后续版本中发生变化。请参考 [CHANGELOG](CHANGELOG.md) 了解版本更新信息。 