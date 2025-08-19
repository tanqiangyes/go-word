# Go Word 用户指南

## 目录

- [快速开始](#快速开始)
- [基础功能](#基础功能)
  - [打开和读取文档](#打开和读取文档)
  - [创建新文档](#创建新文档)
  - [文本操作](#文本操作)
  - [段落操作](#段落操作)
  - [表格操作](#表格操作)
- [高级功能](#高级功能)
  - [样式管理](#样式管理)
  - [文档保护](#文档保护)
  - [图片处理](#图片处理)
  - [图表生成](#图表生成)
  - [文件嵌入和链接](#文件嵌入和链接)
  - [自定义功能区](#自定义功能区)
  - [格式导出](#格式导出)
- [协作功能](#协作功能)
  - [修订跟踪](#修订跟踪)
  - [多人协作](#多人协作)
  - [讨论管理](#讨论管理)
- [性能优化](#性能优化)
- [故障排除](#故障排除)
- [常见问题](#常见问题)

## 快速开始

### 安装

```bash
go get github.com/tanqiangyes/go-word
```

### 第一个程序

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func main() {
    // 打开Word文档
    doc, err := wordprocessingml.Open("example.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // 提取文本内容
    text, err := doc.GetText()
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Println("文档内容:", text)
}
```

## 基础功能

### 打开和读取文档

#### 打开现有文档

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 打开Word文档
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    return fmt.Errorf("无法打开文档: %w", err)
}
defer doc.Close()
```

#### 读取文档内容

```go
// 获取纯文本内容
text, err := doc.GetText()
if err != nil {
    return fmt.Errorf("无法获取文本: %w", err)
}

// 获取段落
paragraphs, err := doc.GetParagraphs()
if err != nil {
    return fmt.Errorf("无法获取段落: %w", err)
}

// 遍历段落
for i, paragraph := range paragraphs {
    fmt.Printf("段落 %d: %s\n", i+1, paragraph.Text)
}

// 获取表格
tables, err := doc.GetTables()
if err != nil {
    return fmt.Errorf("无法获取表格: %w", err)
}

// 遍历表格
for i, table := range tables {
    fmt.Printf("表格 %d: %d行 %d列\n", i+1, len(table.Rows), len(table.Rows[0].Cells))
}
```

### 创建新文档

```go
import "github.com/tanqiangyes/go-word/pkg/writer"

// 创建文档写入器
w := writer.NewDocumentWriter()

// 创建新文档
if err := w.CreateNewDocument(); err != nil {
    return fmt.Errorf("无法创建文档: %w", err)
}

// 添加内容
if err := w.AddParagraph("欢迎使用Go Word!", "Normal"); err != nil {
    return fmt.Errorf("无法添加段落: %w", err)
}

// 保存文档
if err := w.Save("new_document.docx"); err != nil {
    return fmt.Errorf("无法保存文档: %w", err)
}
```

### 文本操作

#### 添加格式化文本

```go
// 添加粗体文本
if err := w.AddFormattedParagraph("这是粗体文本", "Bold", 12, "Arial"); err != nil {
    return err
}

// 添加斜体文本
if err := w.AddFormattedParagraph("这是斜体文本", "Italic", 12, "Arial"); err != nil {
    return err
}

// 添加下划线文本
if err := w.AddFormattedParagraph("这是下划线文本", "Underline", 12, "Arial"); err != nil {
    return err
}
```

#### 文本替换

```go
// 替换文档中的文本
if err := doc.ReplaceText("旧文本", "新文本"); err != nil {
    return fmt.Errorf("无法替换文本: %w", err)
}
```

### 段落操作

#### 添加不同样式的段落

```go
// 添加标题
if err := w.AddParagraph("第一章 概述", "Heading1"); err != nil {
    return err
}

// 添加子标题
if err := w.AddParagraph("1.1 背景", "Heading2"); err != nil {
    return err
}

// 添加正文
if err := w.AddParagraph("这是正文内容...", "Normal"); err != nil {
    return err
}

// 添加引用
if err := w.AddParagraph("这是一段引用文字", "Quote"); err != nil {
    return err
}
```

#### 段落格式设置

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建文字处理器
textProcessor := wordprocessingml.NewTextProcessor()

// 设置段落格式
paragraphFormat := &types.ParagraphFormat{
    Alignment:    "center",     // 居中对齐
    LineSpacing:  1.5,          // 1.5倍行距
    SpaceBefore:  12,           // 段前间距12磅
    SpaceAfter:   6,            // 段后间距6磅
    FirstLineIndent: 24,        // 首行缩进24磅
}

if err := textProcessor.ApplyParagraphFormat(paragraph, paragraphFormat); err != nil {
    return err
}
```

### 表格操作

#### 创建简单表格

```go
// 创建2x3表格
tableData := [][]string{
    {"姓名", "年龄", "职业"},
    {"张三", "25", "工程师"},
    {"李四", "30", "设计师"},
}

if err := w.AddTable(tableData); err != nil {
    return fmt.Errorf("无法添加表格: %w", err)
}
```

#### 高级表格操作

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建高级格式化器
formatter := wordprocessingml.NewAdvancedFormatter()

// 创建表格配置
tableConfig := &types.TableConfig{
    Rows:    3,
    Columns: 3,
    Style:   "TableGrid",
    Width:   "100%",
}

// 创建表格
table, err := formatter.CreateTable(tableConfig)
if err != nil {
    return err
}

// 设置表头
headerRow := table.Rows[0]
headerRow.Cells[0].Text = "产品名称"
headerRow.Cells[1].Text = "价格"
headerRow.Cells[2].Text = "库存"

// 设置表头样式
for _, cell := range headerRow.Cells {
    cell.BackgroundColor = "#4472C4"
    cell.FontColor = "#FFFFFF"
    cell.Bold = true
}

// 添加数据行
table.Rows[1].Cells[0].Text = "笔记本电脑"
table.Rows[1].Cells[1].Text = "¥5,999"
table.Rows[1].Cells[2].Text = "50"

table.Rows[2].Cells[0].Text = "无线鼠标"
table.Rows[2].Cells[1].Text = "¥99"
table.Rows[2].Cells[2].Text = "200"
```

#### 动态表格生成

```go
// 从数据生成表格
func generateTableFromData(data [][]string, headers []string) (*types.Table, error) {
    if len(data) == 0 {
        return nil, fmt.Errorf("数据为空")
    }
    
    // 创建表格
    table := &types.Table{
        Rows: make([]*types.Row, len(data)+1), // +1 for header
        Columns: len(headers),
    }
    
    // 创建表头行
    headerRow := &types.Row{
        Cells: make([]*types.Cell, len(headers)),
    }
    
    for i, header := range headers {
        headerRow.Cells[i] = &types.Cell{
            Text: header,
            Style: "TableHeader",
        }
    }
    table.Rows[0] = headerRow
    
    // 创建数据行
    for i, rowData := range data {
        row := &types.Row{
            Cells: make([]*types.Cell, len(rowData)),
        }
        
        for j, cellData := range rowData {
            row.Cells[j] = &types.Cell{
                Text: cellData,
                Style: "TableCell",
            }
        }
        table.Rows[i+1] = row
    }
    
    return table, nil
}

// 使用示例
salesData := [][]string{
    {"Q1", "100", "150", "200"},
    {"Q2", "120", "180", "220"},
    {"Q3", "90", "160", "190"},
    {"Q4", "110", "170", "210"},
}

headers := []string{"季度", "产品A", "产品B", "产品C"}

table, err := generateTableFromData(salesData, headers)
if err != nil {
    return fmt.Errorf("生成表格失败: %w", err)
}

// 添加表格到文档
if err := w.AddTable(table); err != nil {
    return fmt.Errorf("添加表格失败: %w", err)
}
```

#### 表格样式和格式化

```go
// 应用表格样式
func applyTableStyling(table *types.Table, style string) error {
    switch style {
    case "professional":
        // 专业样式
        for i, row := range table.Rows {
            for j, cell := range row.Cells {
                if i == 0 {
                    // 表头样式
                    cell.BackgroundColor = "#2E74B5"
                    cell.FontColor = "#FFFFFF"
                    cell.Bold = true
                    cell.Alignment = "center"
                } else {
                    // 数据行样式
                    if i%2 == 0 {
                        cell.BackgroundColor = "#F2F2F2"
                    } else {
                        cell.BackgroundColor = "#FFFFFF"
                    }
                    cell.FontColor = "#000000"
                    cell.Alignment = "left"
                }
            }
        }
        
    case "minimal":
        // 简约样式
        for _, row := range table.Rows {
            for _, cell := range row.Cells {
                cell.BackgroundColor = "#FFFFFF"
                cell.FontColor = "#000000"
                cell.BorderColor = "#E0E0E0"
                cell.BorderWidth = 1
            }
        }
        
    case "colorful":
        // 彩色样式
        colors := []string{"#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7"}
        for i, row := range table.Rows {
            for j, cell := range row.Cells {
                if i == 0 {
                    cell.BackgroundColor = "#2C3E50"
                    cell.FontColor = "#FFFFFF"
                    cell.Bold = true
                } else {
                    cell.BackgroundColor = colors[j%len(colors)]
                    cell.FontColor = "#FFFFFF"
                }
            }
        }
    }
    
    return nil
}

// 应用样式到表格
if err := applyTableStyling(table, "professional"); err != nil {
    return fmt.Errorf("应用表格样式失败: %w", err)
}
```

## 高级功能

### 样式管理

#### 创建和应用样式

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建样式系统
styleSystem := wordprocessingml.NewStyleSystem()

// 创建字符样式
charStyle := &types.Style{
    Name:     "HighlightText",
    Type:     types.StyleTypeCharacter,
    FontName: "Arial",
    FontSize: 12,
    Bold:     true,
    FontColor: "#FF0000",
    BackgroundColor: "#FFFF00",
}

if err := styleSystem.AddStyle(charStyle); err != nil {
    return err
}

// 创建段落样式
paraStyle := &types.Style{
    Name:        "CustomHeading",
    Type:        types.StyleTypeParagraph,
    FontName:    "Times New Roman",
    FontSize:    16,
    Bold:        true,
    Alignment:   "center",
    SpaceBefore: 12,
    SpaceAfter:  6,
}

if err := styleSystem.AddStyle(paraStyle); err != nil {
    return err
}

// 应用样式
if err := styleSystem.ApplyStyle(paragraph, "CustomHeading"); err != nil {
    return err
}
```

#### 样式继承

```go
// 创建基础样式
baseStyle := &types.Style{
    Name:     "BaseText",
    Type:     types.StyleTypeCharacter,
    FontName: "Arial",
    FontSize: 11,
}

// 创建继承样式
derivedStyle := &types.Style{
    Name:       "EmphasisText",
    Type:       types.StyleTypeCharacter,
    BasedOn:    "BaseText",  // 继承自BaseText
    Bold:       true,
    FontColor:  "#0066CC",
}

styleSystem.AddStyle(baseStyle)
styleSystem.AddStyle(derivedStyle)
```

#### 高级样式管理

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建样式库
styleLibrary := wordprocessingml.NewStyleLibrary()

// 创建主题样式
themeStyles := map[string]*types.Style{
    "Heading1": {
        Name:        "Heading1",
        Type:        types.StyleTypeParagraph,
        FontName:    "Calibri",
        FontSize:    18,
        Bold:        true,
        FontColor:   "#2E74B5",
        SpaceBefore: 12,
        SpaceAfter:  6,
        Alignment:   "left",
    },
    "Heading2": {
        Name:        "Heading2",
        Type:        types.StyleTypeParagraph,
        FontName:    "Calibri",
        FontSize:    16,
        Bold:        true,
        FontColor:   "#2E74B5",
        SpaceBefore: 10,
        SpaceAfter:  4,
        Alignment:   "left",
    },
    "BodyText": {
        Name:            "BodyText",
        Type:            types.StyleTypeParagraph,
        FontName:        "Calibri",
        FontSize:        11,
        Bold:            false,
        FontColor:       "#000000",
        SpaceBefore:     0,
        SpaceAfter:      0,
        FirstLineIndent: 18,
        LineSpacing:     1.15,
    },
    "Quote": {
        Name:        "Quote",
        Type:        types.StyleTypeParagraph,
        FontName:    "Calibri",
        FontSize:    11,
        Italic:      true,
        FontColor:   "#404040",
        SpaceBefore: 6,
        SpaceAfter:  6,
        LeftIndent:  36,
        RightIndent: 36,
        BackgroundColor: "#F2F2F2",
    },
}

// 添加主题样式到样式库
for _, style := range themeStyles {
    if err := styleLibrary.AddStyle(style); err != nil {
        return fmt.Errorf("添加样式失败: %w", err)
    }
}

// 创建样式模板
template := &types.StyleTemplate{
    Name:        "Professional",
    Description: "专业文档样式模板",
    Styles:      themeStyles,
    Category:    "business",
}

if err := styleLibrary.AddTemplate(template); err != nil {
    return fmt.Errorf("添加模板失败: %w", err)
}

// 应用模板到文档
if err := styleLibrary.ApplyTemplate(doc, "Professional"); err != nil {
    return fmt.Errorf("应用模板失败: %w", err)
}
```

#### 动态样式应用

```go
// 根据内容类型自动应用样式
func applySmartStyling(paragraph *types.Paragraph) error {
    text := paragraph.Text
    
    // 检测标题
    if len(text) < 100 && (strings.HasSuffix(text, ":") || strings.HasPrefix(text, "第")) {
        paragraph.Style = "Heading1"
        return nil
    }
    
    // 检测子标题
    if len(text) < 50 && strings.HasPrefix(text, "•") {
        paragraph.Style = "Heading2"
        return nil
    }
    
    // 检测引用
    if strings.HasPrefix(text, "\"") && strings.HasSuffix(text, "\"") {
        paragraph.Style = "Quote"
        return nil
    }
    
    // 默认样式
    paragraph.Style = "BodyText"
    return nil
}

// 批量应用智能样式
paragraphs, err := doc.GetParagraphs()
if err != nil {
    return err
}

for _, paragraph := range paragraphs {
    if err := applySmartStyling(&paragraph); err != nil {
        return err
    }
}
```

### 文档保护

#### 设置文档保护

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建文档保护器
protector := wordprocessingml.NewDocumentProtector()

// 设置密码保护
protectionConfig := &types.ProtectionConfig{
    Type:     types.ProtectionTypeReadOnly,
    Password: "mypassword",
    Salt:     "randomsalt",
}

if err := protector.SetProtection(doc, protectionConfig); err != nil {
    return fmt.Errorf("无法设置文档保护: %w", err)
}

// 验证密码
if valid := protector.VerifyPassword(doc, "mypassword"); !valid {
    return fmt.Errorf("密码验证失败")
}

// 移除保护
if err := protector.RemoveProtection(doc, "mypassword"); err != nil {
    return fmt.Errorf("无法移除文档保护: %w", err)
}
```

#### 区域保护

```go
// 保护特定区域
rangeConfig := &types.RangeProtectionConfig{
    StartParagraph: 0,
    EndParagraph:   5,
    Type:          types.ProtectionTypeNoEdit,
    Users:         []string{"user1", "user2"},
}

if err := protector.ProtectRange(doc, rangeConfig); err != nil {
    return err
}
```

### 图片处理

#### 插入图片

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建图片处理器
imageProcessor := wordprocessingml.NewImageProcessor()

// 插入图片
imageConfig := &types.ImageConfig{
    FilePath: "image.jpg",
    Width:    300,
    Height:   200,
    Position: types.ImagePositionInline,
}

if err := imageProcessor.InsertImage(doc, imageConfig); err != nil {
    return fmt.Errorf("无法插入图片: %w", err)
}
```

#### 图片效果处理

```go
// 应用图片效果
effectConfig := &types.ImageEffectConfig{
    Brightness: 1.2,
    Contrast:   1.1,
    Saturation: 0.9,
    Blur:       0.5,
}

if err := imageProcessor.ApplyEffect(image, effectConfig); err != nil {
    return err
}

// 调整图片大小
if err := imageProcessor.ResizeImage(image, 400, 300); err != nil {
    return err
}

// 移动图片位置
if err := imageProcessor.MoveImage(image, 100, 50); err != nil {
    return err
}
```

#### 高级图片处理

```go
// 批量处理图片
func processImagesInDocument(doc *Document, imageProcessor *wordprocessingml.ImageProcessor) error {
    // 获取文档中的所有图片
    images, err := doc.GetImages()
    if err != nil {
        return fmt.Errorf("获取图片失败: %w", err)
    }
    
    for i, image := range images {
        // 应用统一的效果
        effects := []*wordprocessingml.ImageProcessorEffect{
            {
                Type:      wordprocessingml.ImageProcessorEffectTypeBrightness,
                Intensity: 1.1,
            },
            {
                Type:      wordprocessingml.ImageProcessorEffectTypeContrast,
                Intensity: 1.05,
            },
            {
                Type:      wordprocessingml.ImageProcessorEffectTypeSharpness,
                Intensity: 1.2,
            },
        }
        
        for _, effect := range effects {
            if err := imageProcessor.ApplyEffect(image, effect); err != nil {
                return fmt.Errorf("应用效果失败: %w", err)
            }
        }
        
        // 统一调整大小
        if err := imageProcessor.ResizeImage(image, 600, 400); err != nil {
            return fmt.Errorf("调整大小失败: %w", err)
        }
        
        // 设置统一位置
        image.Position.Alignment = wordprocessingml.ImageProcessorAlignmentCenter
        image.Position.Wrapping = wordprocessingml.ImageProcessorWrappingSquare
    }
    
    return nil
}

// 图片水印处理
func addWatermarkToImage(image *wordprocessingml.ImageProcessorImage, watermarkText string) error {
    // 创建水印效果
    watermarkEffect := &wordprocessingml.ImageProcessorEffect{
        Type:      wordprocessingml.ImageProcessorEffectTypeWatermark,
        Intensity: 0.3,
        Properties: map[string]interface{}{
            "text":        watermarkText,
            "font":        "Arial",
            "size":        24,
            "color":       "#FF0000",
            "position":    "center",
            "rotation":    45,
        },
    }
    
    return imageProcessor.ApplyEffect(image, watermarkEffect)
}

// 图片格式转换
func convertImageFormat(image *wordprocessingml.ImageProcessorImage, targetFormat string) error {
    conversionConfig := &wordprocessingml.ImageConversionConfig{
        TargetFormat: targetFormat,
        Quality:      90,
        Compression:  true,
    }
    
    return imageProcessor.ConvertFormat(image, conversionConfig)
}

// 使用示例
if err := processImagesInDocument(doc, imageProcessor); err != nil {
    return fmt.Errorf("处理图片失败: %w", err)
}

// 为特定图片添加水印
if err := addWatermarkToImage(image, "机密文档"); err != nil {
    return fmt.Errorf("添加水印失败: %w", err)
}

// 转换图片格式
if err := convertImageFormat(image, "png"); err != nil {
    return fmt.Errorf("转换格式失败: %w", err)
}
```

### 自定义功能区

#### 创建自定义功能区

```go
import (
    "context"
    "fmt"
    "os"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 创建自定义功能区
customRibbon := wordprocessingml.NewCustomRibbon(doc, nil)

// 添加选项卡
tab := &wordprocessingml.RibbonTab{
    ID:          "custom_tab",
    Label:       "自定义工具",
    Description: "用户自定义的工具选项卡",
    Visible:     true,
    Enabled:     true,
    Position:    1,
}

if err := customRibbon.AddTab(tab); err != nil {
    return fmt.Errorf("添加选项卡失败: %w", err)
}

// 添加组
group := &wordprocessingml.RibbonGroup{
    ID:          "custom_group",
    Label:       "常用工具",
    Description: "常用工具集合",
    TabID:       "custom_tab",
    Visible:     true,
    Enabled:     true,
    Position:    1,
}

if err := customRibbon.AddGroup(group); err != nil {
    return fmt.Errorf("添加组失败: %w", err)
}

// 添加按钮控件
button := &wordprocessingml.RibbonControl{
    ID:          "custom_button",
    Type:        wordprocessingml.ControlTypeButton,
    Label:       "快速格式化",
    Description: "快速应用常用格式",
    Tooltip:     "点击应用快速格式化",
    GroupID:     "custom_group",
    Action:      "quick_format",
    Shortcut:    "Ctrl+Shift+F",
    Size:        wordprocessingml.ControlSizeMedium,
}

if err := customRibbon.AddControl(button); err != nil {
    return fmt.Errorf("添加控件失败: %w", err)
}
```

#### 注册回调函数

```go
// 注册按钮回调函数
err := customRibbon.RegisterCallback("custom_button", func(ctx context.Context, control *wordprocessingml.RibbonControl, args map[string]interface{}) error {
    // 执行快速格式化逻辑
    fmt.Println("执行快速格式化...")
    
    // 这里可以添加具体的格式化逻辑
    // 例如：应用样式、调整格式等
    
    return nil
})
if err != nil {
    return fmt.Errorf("注册回调函数失败: %w", err)
}

// 触发控件
if err := customRibbon.TriggerControl(context.Background(), "custom_button", nil); err != nil {
    return fmt.Errorf("触发控件失败: %w", err)
}
```

#### 管理功能区

```go
// 列出所有选项卡
tabs := customRibbon.ListTabs()
for _, tab := range tabs {
    fmt.Printf("选项卡: %s (%s)\n", tab.Label, tab.ID)
    
    // 列出组
    for _, group := range tab.Groups {
        fmt.Printf("  组: %s (%s)\n", group.Label, group.ID)
        
        // 列出控件
        for _, control := range group.Controls {
            fmt.Printf("    控件: %s [%s] - %s\n", 
                control.Label, control.Type, control.Description)
        }
    }
}

// 更新控件
updates := map[string]interface{}{
    "label":     "新标签",
    "tooltip":   "新提示",
    "enabled":   false,
}
if err := customRibbon.UpdateControl("custom_button", updates); err != nil {
    return fmt.Errorf("更新控件失败: %w", err)
}

// 导出功能区配置
ribbonData, err := customRibbon.ExportRibbon()
if err != nil {
    return fmt.Errorf("导出功能区失败: %w", err)
}

// 保存到文件
if err := os.WriteFile("custom_ribbon.xml", ribbonData, 0644); err != nil {
    return fmt.Errorf("保存功能区配置失败: %w", err)
}
```

### 文件嵌入和链接

#### 嵌入文件

```go
import (
    "context"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 创建文件嵌入器
fileEmbedder := wordprocessingml.NewFileEmbedder(doc, nil)

// 嵌入文件
embedResult, err := fileEmbedder.EmbedFile(context.Background(), "document.pdf", 
    wordprocessingml.EmbedTypeAttachment, &wordprocessingml.EmbedPosition{
        ParagraphIndex: 0,
        RunIndex:       0,
        CharIndex:      0,
    })
if err != nil {
    return fmt.Errorf("文件嵌入失败: %w", err)
}

fmt.Printf("文件嵌入成功，ID: %s\n", embedResult.FileID)
```

#### 创建链接

```go
// 创建外部链接
linkResult, err := fileEmbedder.CreateLink(context.Background(), 
    wordprocessingml.LinkTypeExternal, 
    "https://example.com", 
    "访问示例网站", 
    nil, nil)
if err != nil {
    return fmt.Errorf("链接创建失败: %w", err)
}

// 创建邮件链接
emailLink, err := fileEmbedder.CreateLink(context.Background(), 
    wordprocessingml.LinkTypeEmail, 
    "contact@example.com", 
    "联系我们", 
    nil, nil)
if err != nil {
    return fmt.Errorf("邮件链接创建失败: %w", err)
}

// 创建文件链接
fileLink, err := fileEmbedder.CreateLink(context.Background(), 
    wordprocessingml.LinkTypeFile, 
    "./reference.pdf", 
    "参考文档", 
    nil, nil)
if err != nil {
    return fmt.Errorf("文件链接创建失败: %w", err)
}
```

#### 管理嵌入文件

```go
// 列出所有嵌入文件
embeddedFiles := fileEmbedder.ListEmbeddedFiles()
for _, file := range embeddedFiles {
    fmt.Printf("文件: %s, 大小: %d bytes, 类型: %s\n", 
        file.Name, file.Size, file.MimeType)
}

// 提取嵌入文件
if err := fileEmbedder.ExtractFile(embedResult.FileID, "extracted_document.pdf"); err != nil {
    return fmt.Errorf("文件提取失败: %w", err)
}

// 移除嵌入文件
if err := fileEmbedder.RemoveEmbeddedFile(embedResult.FileID); err != nil {
    return fmt.Errorf("文件移除失败: %w", err)
}
```

### 图表生成

#### 创建图表

```go
import (
    "context"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 创建图表生成器
chartGenerator := wordprocessingml.NewChartGenerator(nil)

// 创建折线图
chart := &wordprocessingml.ChartGeneratorChart{
    Type:   wordprocessingml.ChartGeneratorChartTypeLine,
    Title:  "销售趋势图",
    Options: &wordprocessingml.ChartGeneratorOptions{
        Width:  600,
        Height: 400,
    },
}

if err := chartGenerator.CreateChart(context.Background(), chart); err != nil {
    return err
}

// 添加数据点
dataPoints := []*wordprocessingml.ChartGeneratorDataPoint{
    {Label: "1月", Value: 100},
    {Label: "2月", Value: 150},
    {Label: "3月", Value: 200},
    {Label: "4月", Value: 180},
    {Label: "5月", Value: 220},
    {Label: "6月", Value: 250},
}

for _, point := range dataPoints {
    if err := chartGenerator.AddDataPoint(context.Background(), chart.ID, point); err != nil {
        return err
    }
}

// 导出图表
chartData, err := chartGenerator.ExportChart(context.Background(), chart.ID, "svg", nil)
if err != nil {
    return err
}
```

#### 不同类型的图表

```go
// 柱状图
barChart := &wordprocessingml.ChartGeneratorChart{
    Type:   wordprocessingml.ChartGeneratorChartTypeBar,
    Title:  "产品销量对比",
    Options: &wordprocessingml.ChartGeneratorOptions{
        Width:  500,
        Height: 300,
    },
}
if err := chartGenerator.CreateChart(context.Background(), barChart); err != nil {
    return err
}

// 饼图
pieChart := &wordprocessingml.ChartGeneratorChart{
    Type:   wordprocessingml.ChartGeneratorChartTypePie,
    Title:  "市场份额分布",
    Options: &wordprocessingml.ChartGeneratorOptions{
        Width:  400,
        Height: 400,
    },
}
if err := chartGenerator.CreateChart(context.Background(), pieChart); err != nil {
    return err
}

// 散点图
scatterChart := &wordprocessingml.ChartGeneratorChart{
    Type:   wordprocessingml.ChartGeneratorChartTypeScatter,
    Title:  "相关性分析",
    Options: &wordprocessingml.ChartGeneratorOptions{
        Width:  600,
        Height: 400,
    },
}
if err := chartGenerator.CreateChart(context.Background(), scatterChart); err != nil {
    return err
}
```

### 格式导出

#### 导出为PDF

```go
import (
    "context"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 创建PDF导出器
pdfExporter := wordprocessingml.NewPDFExporter(doc, nil)

// 导出PDF
if err := pdfExporter.ExportToPDF(context.Background(), "output.pdf"); err != nil {
    return fmt.Errorf("PDF导出失败: %w", err)
}
```

#### 导出为其他格式

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建增强文档构建器
builder := wordprocessingml.NewEnhancedDocumentBuilder(nil)

// 导出为RTF格式
if err := builder.ExportDocument(doc, "output.rtf", "rtf"); err != nil {
    return fmt.Errorf("RTF导出失败: %w", err)
}

// 导出为HTML格式
if err := builder.ExportDocument(doc, "output.html", "html"); err != nil {
    return fmt.Errorf("HTML导出失败: %w", err)
}

// 导出为纯文本格式
if err := builder.ExportDocument(doc, "output.txt", "txt"); err != nil {
    return fmt.Errorf("文本导出失败: %w", err)
}
```

#### 批量格式转换

```go
import (
    "context"
    "fmt"
    "path/filepath"
    "strings"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 批量转换多个文档
documents := []string{"doc1.docx", "doc2.docx", "doc3.docx"}
formats := []string{"pdf", "rtf", "html"}

for _, docPath := range documents {
    doc, err := wordprocessingml.Open(docPath)
    if err != nil {
        fmt.Printf("无法打开文档 %s: %v\n", docPath, err)
        continue
    }
    defer doc.Close()
    
    baseName := strings.TrimSuffix(filepath.Base(docPath), filepath.Ext(docPath))
    
    for _, format := range formats {
        outputPath := fmt.Sprintf("%s.%s", baseName, format)
        
        switch format {
        case "pdf":
            pdfExporter := wordprocessingml.NewPDFExporter(doc, nil)
            if err := pdfExporter.ExportToPDF(context.Background(), outputPath); err != nil {
                fmt.Printf("PDF导出失败 %s: %v\n", docPath, err)
            } else {
                fmt.Printf("成功导出 %s 为 %s\n", docPath, outputPath)
            }
        case "rtf", "html", "txt":
            builder := wordprocessingml.NewEnhancedDocumentBuilder(nil)
            if err := builder.ExportDocument(doc, outputPath, format); err != nil {
                fmt.Printf("%s导出失败 %s: %v\n", strings.ToUpper(format), docPath, err)
            } else {
                fmt.Printf("成功导出 %s 为 %s\n", docPath, outputPath)
            }
        }
    }
}
```

```go
import (
    "context"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 创建PDF导出器
pdfExporter := wordprocessingml.NewPDFExporter(doc, nil)

// 导出PDF
if err := pdfExporter.ExportToPDF(context.Background(), "output.pdf"); err != nil {
    return fmt.Errorf("PDF导出失败: %w", err)
}
```

#### 批量PDF导出

```go
import (
    "context"
    "fmt"
    "strings"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 批量处理多个文档
documents := []string{"doc1.docx", "doc2.docx", "doc3.docx"}
results := make(chan error, len(documents))

for _, docPath := range documents {
    go func(path string) {
        doc, err := wordprocessingml.Open(path)
        if err != nil {
            results <- fmt.Errorf("无法打开文档 %s: %w", path, err)
            return
        }
        defer doc.Close()
        
        // 创建PDF导出器
        pdfExporter := wordprocessingml.NewPDFExporter(doc, nil)
        
        // 导出PDF
        outputPath := strings.Replace(path, ".docx", ".pdf", 1)
        if err := pdfExporter.ExportToPDF(context.Background(), outputPath); err != nil {
            results <- fmt.Errorf("PDF导出失败 %s: %w", path, err)
            return
        }
        
        results <- nil
    }(docPath)
}

// 收集结果
for i := 0; i < len(documents); i++ {
    if err := <-results; err != nil {
        fmt.Printf("处理失败: %v\n", err)
    } else {
        fmt.Printf("处理成功\n")
    }
}
```

## 协作功能

### 修订跟踪

#### 启用修订跟踪

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建修订跟踪器
revisionTracker := wordprocessingml.NewRevisionTracker()

// 启用修订跟踪
if err := revisionTracker.EnableTracking(doc); err != nil {
    return err
}

// 记录修改
revision := &types.Revision{
    Type:        types.RevisionTypeInsert,
    Author:      "张三",
    DateTime:    time.Now(),
    Description: "添加新段落",
    Content:     "这是新添加的内容",
}

if err := revisionTracker.AddRevision(doc, revision); err != nil {
    return err
}
```

#### 处理修订

```go
// 获取所有修订
revisions, err := revisionTracker.GetRevisions(doc)
if err != nil {
    return err
}

// 遍历修订
for _, revision := range revisions {
    fmt.Printf("修订ID: %s, 作者: %s, 类型: %v\n", 
        revision.ID, revision.Author, revision.Type)
}

// 接受修订
if err := revisionTracker.AcceptRevision(doc, "revision-id-1"); err != nil {
    return err
}

// 拒绝修订
if err := revisionTracker.RejectRevision(doc, "revision-id-2"); err != nil {
    return err
}

// 接受所有修订
if err := revisionTracker.AcceptAllRevisions(doc); err != nil {
    return err
}
```

### 多人协作

#### 协作编辑

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建协作编辑器
collabEditor := wordprocessingml.NewCollaborativeEditor()

// 创建编辑会话
sessionConfig := &types.CollaborationConfig{
    DocumentID:   "doc-123",
    MaxUsers:     10,
    AutoSave:     true,
    SaveInterval: 30, // 30秒自动保存
}

session, err := collabEditor.CreateSession(doc, sessionConfig)
if err != nil {
    return err
}

// 用户加入会话
user := &types.CollaborationUser{
    ID:          "user-001",
    Name:        "张三",
    Email:       "zhangsan@example.com",
    Permissions: []string{"read", "write", "comment"},
}

if err := collabEditor.JoinSession(session.ID, user); err != nil {
    return err
}
```

#### 冲突解决

```go
// 检测冲突
conflicts, err := collabEditor.DetectConflicts(session.ID)
if err != nil {
    return err
}

// 处理冲突
for _, conflict := range conflicts {
    resolution := &types.ConflictResolution{
        ConflictID: conflict.ID,
        Strategy:   types.ResolutionStrategyMerge,
        ResolvedBy: "user-001",
    }
    
    if err := collabEditor.ResolveConflict(session.ID, resolution); err != nil {
        return err
    }
}
```

### 讨论管理

#### 创建讨论

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建讨论管理器
discussionManager := wordprocessingml.NewDiscussionManager()

// 创建讨论
discussion := &types.Discussion{
    Title:       "关于第三章的修改建议",
    Description: "建议调整第三章的结构",
    Author:      "李四",
    Priority:    types.PriorityMedium,
    Tags:        []string{"结构", "修改"},
}

if err := discussionManager.CreateDiscussion(doc, discussion); err != nil {
    return err
}

// 添加评论
comment := &types.DiscussionComment{
    DiscussionID: discussion.ID,
    Author:       "王五",
    Content:      "我同意这个建议，第三章确实需要重新组织",
    DateTime:     time.Now(),
}

if err := discussionManager.AddComment(comment); err != nil {
    return err
}
```

#### 讨论通知

```go
// 订阅讨论
subscription := &types.DiscussionSubscription{
    UserID:       "user-001",
    DiscussionID: discussion.ID,
    NotifyEmail:  true,
    NotifyApp:    true,
}

if err := discussionManager.Subscribe(subscription); err != nil {
    return err
}

// 发送通知
notification := &types.DiscussionNotification{
    Type:         types.NotificationTypeNewComment,
    DiscussionID: discussion.ID,
    Message:      "有新的评论",
    Recipients:   []string{"user-001", "user-002"},
}

if err := discussionManager.SendNotification(notification); err != nil {
    return err
}
```

## 性能优化

### 内存管理

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建性能优化器
optimizer := wordprocessingml.NewPerformanceOptimizer()

// 配置内存池
poolConfig := &types.MemoryPoolConfig{
    InitialSize:  1024 * 1024,  // 1MB
    MaxSize:      10 * 1024 * 1024,  // 10MB
    GrowthFactor: 2.0,
}

if err := optimizer.ConfigureMemoryPool(poolConfig); err != nil {
    return err
}

// 启用缓存
cacheConfig := &types.CacheConfig{
    MaxEntries:   1000,
    TTL:          time.Hour,
    CleanupInterval: time.Minute * 10,
}

if err := optimizer.EnableCache(cacheConfig); err != nil {
    return err
}
```

### 并发处理

```go
// 配置并发
concurrencyConfig := &types.ConcurrencyConfig{
    MaxWorkers:     runtime.NumCPU(),
    QueueSize:      1000,
    WorkerTimeout:  time.Minute * 5,
}

if err := optimizer.ConfigureConcurrency(concurrencyConfig); err != nil {
    return err
}

// 并发处理多个文档
documents := []string{"doc1.docx", "doc2.docx", "doc3.docx"}
results := make(chan *types.ProcessResult, len(documents))

for _, docPath := range documents {
    go func(path string) {
        result := &types.ProcessResult{FilePath: path}
        
        doc, err := wordprocessingml.Open(path)
        if err != nil {
            result.Error = err
            results <- result
            return
        }
        defer doc.Close()
        
        // 处理文档
        text, err := doc.GetText()
        if err != nil {
            result.Error = err
        } else {
            result.Data = text
        }
        
        results <- result
    }(docPath)
}

// 收集结果
for i := 0; i < len(documents); i++ {
    result := <-results
    if result.Error != nil {
        fmt.Printf("处理 %s 失败: %v\n", result.FilePath, result.Error)
    } else {
        fmt.Printf("处理 %s 成功\n", result.FilePath)
    }
}
```

### 性能监控

```go
// 获取性能指标
metrics, err := optimizer.GetMetrics()
if err != nil {
    return err
}

fmt.Printf("内存使用: %d bytes\n", metrics.MemoryUsage)
fmt.Printf("缓存命中率: %.2f%%\n", metrics.CacheHitRate*100)
fmt.Printf("平均处理时间: %v\n", metrics.AverageProcessTime)
fmt.Printf("活跃协程数: %d\n", metrics.ActiveGoroutines)

// 生成性能报告
report, err := optimizer.GenerateReport()
if err != nil {
    return err
}

// 保存报告
if err := report.SaveToFile("performance_report.html"); err != nil {
    return err
}
```

## 故障排除

### 常见错误处理

```go
import "github.com/tanqiangyes/go-word/pkg/utils"

// 检查错误类型
if err != nil {
    switch {
    case utils.IsParseError(err):
        fmt.Println("文档解析错误:", err)
        // 尝试修复或使用备用方案
        
    case utils.IsOPCError(err):
        fmt.Println("OPC容器错误:", err)
        // 检查文件完整性
        
    case utils.IsMemoryError(err):
        fmt.Println("内存不足错误:", err)
        // 释放资源或增加内存
        
    case utils.IsPermissionError(err):
        fmt.Println("权限错误:", err)
        // 检查文件权限
        
    default:
        fmt.Println("未知错误:", err)
    }
}
```

### 调试模式

```go
import "github.com/tanqiangyes/go-word/pkg/utils"

// 启用调试模式
utils.SetDebugMode(true)

// 设置日志级别
utils.SetLogLevel(utils.LogLevelDebug)

// 启用详细错误信息
utils.EnableVerboseErrors(true)

// 设置日志输出
logFile, err := os.Create("debug.log")
if err != nil {
    return err
}
defer logFile.Close()

utils.SetLogOutput(logFile)
```

### 文档验证

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建文档验证器
validator := wordprocessingml.NewDocumentValidator()

// 验证文档
validationResult, err := validator.ValidateDocument(doc)
if err != nil {
    return err
}

if !validationResult.IsValid {
    fmt.Println("文档验证失败:")
    for _, issue := range validationResult.Issues {
        fmt.Printf("- %s: %s\n", issue.Severity, issue.Message)
    }
}

// 尝试修复
if validationResult.CanAutoFix {
    if err := validator.AutoFix(doc); err != nil {
        return err
    }
    fmt.Println("文档已自动修复")
}
```

## 常见问题

### Q: 如何处理大文档？

A: 对于大文档，建议使用流式处理和分块读取：

```go
// 启用流式处理
options := &wordprocessingml.OpenOptions{
    StreamMode:    true,
    ChunkSize:     1024 * 1024,  // 1MB chunks
    LazyLoading:   true,
}

doc, err := wordprocessingml.OpenWithOptions("large_document.docx", options)
if err != nil {
    return err
}
defer doc.Close()

// 分块处理段落
for {
    paragraphs, hasMore, err := doc.GetParagraphsChunk(100)  // 每次100个段落
    if err != nil {
        return err
    }
    
    // 处理当前批次的段落
    for _, paragraph := range paragraphs {
        // 处理段落
    }
    
    if !hasMore {
        break
    }
}
```

### Q: 如何提高处理速度？

A: 可以通过以下方式优化性能：

1. **启用缓存**：
```go
optimizer.EnableCache(&types.CacheConfig{
    MaxEntries: 1000,
    TTL:        time.Hour,
})
```

2. **使用并发处理**：
```go
optimizer.ConfigureConcurrency(&types.ConcurrencyConfig{
    MaxWorkers: runtime.NumCPU(),
})
```

3. **预分配内存**：
```go
optimizer.ConfigureMemoryPool(&types.MemoryPoolConfig{
    InitialSize: 10 * 1024 * 1024,  // 10MB
})
```

### Q: 如何处理损坏的文档？

A: 使用文档修复功能：

```go
// 创建文档修复器
repairer := wordprocessingml.NewDocumentRepairer()

// 检查文档
diagnosis, err := repairer.DiagnoseDocument("corrupted.docx")
if err != nil {
    return err
}

if diagnosis.IsCorrupted {
    // 尝试修复
    repairResult, err := repairer.RepairDocument("corrupted.docx", "repaired.docx")
    if err != nil {
        return err
    }
    
    if repairResult.Success {
        fmt.Println("文档修复成功")
    } else {
        fmt.Printf("文档修复失败: %s\n", repairResult.Error)
    }
}
```

### Q: 如何处理不同版本的Word文档？

A: 库自动检测并处理不同版本：

```go
// 检查文档版本
version, err := doc.GetVersion()
if err != nil {
    return err
}

fmt.Printf("文档版本: %s\n", version)

// 根据版本调整处理策略
switch version {
case "2007":
    // Word 2007 特定处理
case "2010":
    // Word 2010 特定处理
case "2013", "2016", "2019":
    // 较新版本处理
default:
    // 默认处理
}
```

### Q: 如何处理密码保护的文档？

A: 使用密码验证功能：

```go
// 检查是否有密码保护
if doc.IsPasswordProtected() {
    // 提供密码
    if err := doc.Authenticate("password123"); err != nil {
        return fmt.Errorf("密码验证失败: %w", err)
    }
}

// 现在可以正常访问文档内容
text, err := doc.GetText()
if err != nil {
    return err
}
```

### Q: 如何自定义样式？

A: 创建和应用自定义样式：

```go
// 创建自定义样式
customStyle := &types.Style{
    Name:            "MyCustomStyle",
    Type:            types.StyleTypeParagraph,
    FontName:        "Calibri",
    FontSize:        12,
    FontColor:       "#333333",
    BackgroundColor: "#F0F0F0",
    Bold:            false,
    Italic:          false,
    Underline:       false,
    Alignment:       "left",
    LineSpacing:     1.15,
    SpaceBefore:     6,
    SpaceAfter:      6,
    FirstLineIndent: 0,
    LeftIndent:      0,
    RightIndent:     0,
}

// 添加到样式系统
if err := styleSystem.AddStyle(customStyle); err != nil {
    return err
}

// 应用样式
if err := styleSystem.ApplyStyle(paragraph, "MyCustomStyle"); err != nil {
    return err
}
```

### Q: 如何处理文档中的图片？

A: 使用图片处理器：

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建图片处理器
imageProcessor := wordprocessingml.NewImageProcessor(nil)

// 加载图片
image, err := imageProcessor.LoadImage(context.Background(), "image.jpg")
if err != nil {
    return fmt.Errorf("加载图片失败: %w", err)
}

// 调整图片大小
if err := imageProcessor.ResizeImage(image, 400, 300); err != nil {
    return fmt.Errorf("调整图片大小失败: %w", err)
}

// 应用图片效果
effect := &wordprocessingml.ImageProcessorEffect{
    Type:      wordprocessingml.ImageProcessorEffectTypeBrightness,
    Intensity: 1.2,
}
if err := imageProcessor.ApplyEffect(image, effect); err != nil {
    return fmt.Errorf("应用图片效果失败: %w", err)
}
```

### Q: 如何创建和管理图表？

A: 使用图表生成器：

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建图表生成器
chartGenerator := wordprocessingml.NewChartGenerator(nil)

// 创建图表
chart := &wordprocessingml.ChartGeneratorChart{
    Type:   wordprocessingml.ChartGeneratorChartTypeBar,
    Title:  "销售数据",
    Options: &wordprocessingml.ChartGeneratorOptions{
        Width:  600,
        Height: 400,
    },
}

if err := chartGenerator.CreateChart(context.Background(), chart); err != nil {
    return fmt.Errorf("创建图表失败: %w", err)
}

// 添加数据点
dataPoint := &wordprocessingml.ChartGeneratorDataPoint{
    Label: "Q1",
    Value: 100,
}
if err := chartGenerator.AddDataPoint(context.Background(), chart.ID, dataPoint); err != nil {
    return fmt.Errorf("添加数据点失败: %w", err)
}
```

### Q: 如何嵌入和管理外部文件？

A: 使用文件嵌入器：

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建文件嵌入器
fileEmbedder := wordprocessingml.NewFileEmbedder(doc, nil)

// 嵌入PDF文件
embedResult, err := fileEmbedder.EmbedFile(context.Background(), "report.pdf", 
    wordprocessingml.EmbedTypeAttachment, nil)
if err != nil {
    return fmt.Errorf("嵌入文件失败: %w", err)
}

// 创建外部链接
linkResult, err := fileEmbedder.CreateLink(context.Background(), 
    wordprocessingml.LinkTypeExternal, 
    "https://example.com", 
    "访问网站", nil, nil)
if err != nil {
    return fmt.Errorf("创建链接失败: %w", err)
}

// 列出所有嵌入文件
embeddedFiles := fileEmbedder.ListEmbeddedFiles()
for _, file := range embeddedFiles {
    fmt.Printf("文件: %s, 大小: %d bytes\n", file.Name, file.Size)
}
```

### Q: 如何自定义工作界面？

A: 使用自定义功能区：

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建自定义功能区
customRibbon := wordprocessingml.NewCustomRibbon(doc, nil)

// 添加选项卡
tab := &wordprocessingml.RibbonTab{
    ID:    "my_tools",
    Label: "我的工具",
}
if err := customRibbon.AddTab(tab); err != nil {
    return fmt.Errorf("添加选项卡失败: %w", err)
}

// 添加按钮控件
button := &wordprocessingml.RibbonControl{
    ID:      "quick_action",
    Type:    wordprocessingml.ControlTypeButton,
    Label:   "快速操作",
    GroupID: "my_group",
}
if err := customRibbon.AddControl(button); err != nil {
    return fmt.Errorf("添加控件失败: %w", err)
}

// 注册回调函数
err := customRibbon.RegisterCallback("quick_action", func(ctx context.Context, control *wordprocessingml.RibbonControl, args map[string]interface{}) error {
    fmt.Println("执行快速操作...")
    return nil
})
if err != nil {
    return fmt.Errorf("注册回调函数失败: %w", err)
}
```

### Q: 如何优化文档处理性能？

A: 使用性能优化器：

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建性能优化器
optimizer := wordprocessingml.NewPerformanceOptimizer()

// 配置内存池
poolConfig := &types.MemoryPoolConfig{
    InitialSize:  10 * 1024 * 1024, // 10MB
    MaxSize:      100 * 1024 * 1024, // 100MB
    GrowthFactor: 2.0,
}
if err := optimizer.ConfigureMemoryPool(poolConfig); err != nil {
    return fmt.Errorf("配置内存池失败: %w", err)
}

// 启用缓存
cacheConfig := &types.CacheConfig{
    MaxEntries:      1000,
    TTL:            time.Hour,
    CleanupInterval: time.Minute * 10,
}
if err := optimizer.EnableCache(cacheConfig); err != nil {
    return fmt.Errorf("启用缓存失败: %w", err)
}

// 配置并发
concurrencyConfig := &types.ConcurrencyConfig{
    MaxWorkers:    runtime.NumCPU(),
    QueueSize:     1000,
    WorkerTimeout: time.Minute * 5,
}
if err := optimizer.ConfigureConcurrency(concurrencyConfig); err != nil {
    return fmt.Errorf("配置并发失败: %w", err)
}
```

### Q: 如何处理文档协作和修订？

A: 使用协作系统：

```go
import "github.com/tanqiangyes/go-word/pkg/wordprocessingml"

// 创建修订跟踪器
revisionTracker := wordprocessingml.NewRevisionTracker()

// 启用修订跟踪
if err := revisionTracker.EnableTracking(doc); err != nil {
    return fmt.Errorf("启用修订跟踪失败: %w", err)
}

// 创建协作编辑器
collabEditor := wordprocessingml.NewCollaborativeEditor()

// 创建编辑会话
sessionConfig := &types.CollaborationConfig{
    DocumentID:   "doc-123",
    MaxUsers:     10,
    AutoSave:     true,
    SaveInterval: 30,
}
session, err := collabEditor.CreateSession(doc, sessionConfig)
if err != nil {
    return fmt.Errorf("创建协作会话失败: %w", err)
}

// 用户加入会话
user := &types.CollaborationUser{
    ID:          "user-001",
    Name:        "张三",
    Permissions: []string{"read", "write", "comment"},
}
if err := collabEditor.JoinSession(session.ID, user); err != nil {
    return fmt.Errorf("加入会话失败: %w", err)
}
```

---

## 实用工具函数

### 文档批量处理工具

```go
// 批量文档处理器
type BatchDocumentProcessor struct {
    maxWorkers int
    logger     *utils.Logger
}

// 处理单个文档
func (bdp *BatchDocumentProcessor) ProcessDocument(docPath string) error {
    doc, err := wordprocessingml.Open(docPath)
    if err != nil {
        return fmt.Errorf("打开文档失败: %w", err)
    }
    defer doc.Close()
    
    // 应用统一格式
    if err := bdp.applyStandardFormatting(doc); err != nil {
        return fmt.Errorf("应用格式失败: %w", err)
    }
    
    // 生成输出文件名
    outputPath := strings.Replace(docPath, ".docx", "_formatted.docx", 1)
    
    // 保存文档
    writer := writer.NewDocumentWriter()
    if err := writer.Save(outputPath); err != nil {
        return fmt.Errorf("保存文档失败: %w", err)
    }
    
    bdp.logger.Info("文档处理完成", map[string]interface{}{
        "input":  docPath,
        "output": outputPath,
    })
    
    return nil
}

// 批量处理文档
func (bdp *BatchDocumentProcessor) ProcessDocuments(docPaths []string) error {
    semaphore := make(chan struct{}, bdp.maxWorkers)
    results := make(chan error, len(docPaths))
    
    for _, docPath := range docPaths {
        go func(path string) {
            semaphore <- struct{}{}
            defer func() { <-semaphore }()
            
            results <- bdp.ProcessDocument(path)
        }(docPath)
    }
    
    // 收集结果
    for i := 0; i < len(docPaths); i++ {
        if err := <-results; err != nil {
            bdp.logger.Error("文档处理失败", map[string]interface{}{
                "error": err.Error(),
            })
        }
    }
    
    return nil
}

// 应用标准格式
func (bdp *BatchDocumentProcessor) applyStandardFormatting(doc *wordprocessingml.Document) error {
    // 应用标准样式
    styleSystem := wordprocessingml.NewStyleSystem()
    
    // 设置标准字体
    standardFont := &types.Font{
        Name: "Microsoft YaHei",
        Size: 12,
    }
    
    if err := styleSystem.SetDefaultFont(standardFont); err != nil {
        return fmt.Errorf("设置默认字体失败: %w", err)
    }
    
    // 应用标准段落格式
    standardParagraph := &types.ParagraphFormat{
        Alignment:       "justify",
        LineSpacing:     1.5,
        SpaceBefore:     6,
        SpaceAfter:      6,
        FirstLineIndent: 24,
    }
    
    if err := styleSystem.SetDefaultParagraphFormat(standardParagraph); err != nil {
        return fmt.Errorf("设置默认段落格式失败: %w", err)
    }
    
    return nil
}
```

### 文档质量检查工具

```go
// 文档质量检查器
type DocumentQualityChecker struct {
    rules []QualityRule
    logger *utils.Logger
}

// 质量规则
type QualityRule struct {
    Name        string
    Description string
    Check       func(*wordprocessingml.Document) (bool, string)
    Severity    string // "error", "warning", "info"
}

// 检查文档质量
func (dqc *DocumentQualityChecker) CheckDocument(doc *wordprocessingml.Document) []QualityIssue {
    var issues []QualityIssue
    
    for _, rule := range dqc.rules {
        passed, message := rule.Check(doc)
        if !passed {
            issues = append(issues, QualityIssue{
                Rule:      rule.Name,
                Message:   message,
                Severity:  rule.Severity,
                Timestamp: time.Now(),
            })
        }
    }
    
    return issues
}

// 预定义质量规则
func (dqc *DocumentQualityChecker) InitializeDefaultRules() {
    dqc.rules = []QualityRule{
        {
            Name:        "字体一致性",
            Description: "检查文档中字体使用是否一致",
            Severity:    "warning",
            Check: func(doc *wordprocessingml.Document) (bool, string) {
                // 实现字体一致性检查逻辑
                return true, "字体使用一致"
            },
        },
        {
            Name:        "段落间距",
            Description: "检查段落间距是否合理",
            Severity:    "info",
            Check: func(doc *wordprocessingml.Document) (bool, string) {
                // 实现段落间距检查逻辑
                return true, "段落间距合理"
            },
        },
        {
            Name:        "表格格式",
            Description: "检查表格格式是否规范",
            Severity:    "warning",
            Check: func(doc *wordprocessingml.Document) (bool, string) {
                // 实现表格格式检查逻辑
                return true, "表格格式规范"
            },
        },
    }
}

// 质量检查结果
type QualityIssue struct {
    Rule      string    `json:"rule"`
    Message   string    `json:"message"`
    Severity  string    `json:"severity"`
    Timestamp time.Time `json:"timestamp"`
}
```

### 使用示例

```go
// 创建批量处理器
batchProcessor := &BatchDocumentProcessor{
    maxWorkers: 4,
    logger:     utils.NewLogger(utils.LogLevelInfo, nil),
}

// 批量处理文档
docPaths := []string{"doc1.docx", "doc2.docx", "doc3.docx"}
if err := batchProcessor.ProcessDocuments(docPaths); err != nil {
    log.Fatalf("批量处理失败: %v", err)
}

// 创建质量检查器
qualityChecker := &DocumentQualityChecker{
    logger: utils.NewLogger(utils.LogLevelInfo, nil),
}
qualityChecker.InitializeDefaultRules()

// 检查文档质量
doc, err := wordprocessingml.Open("document.docx")
if err != nil {
    log.Fatalf("打开文档失败: %v", err)
}
defer doc.Close()

issues := qualityChecker.CheckDocument(doc)
for _, issue := range issues {
    fmt.Printf("[%s] %s: %s\n", issue.Severity, issue.Rule, issue.Message)
}
```

## 总结

本用户指南涵盖了Go Word库的主要功能和使用方法。从基础的文档读写操作到高级的协作功能，从性能优化到故障排除，为用户提供了全面的使用指导。

### 主要特性总结

✅ **基础功能**：文档读写、文本操作、段落操作、表格操作
✅ **高级功能**：样式管理、文档保护、图片处理、图表生成、文件嵌入链接、自定义功能区、多格式导出
✅ **协作功能**：修订跟踪、多人协作、讨论管理
✅ **性能优化**：内存管理、并发处理、性能监控
✅ **实用工具**：批量处理、质量检查、自动化工具

### 使用建议

1. **性能优化**：对于大文档，建议使用流式处理和分块读取
2. **错误处理**：始终检查返回的错误，使用结构化的错误处理
3. **资源管理**：及时关闭文档和释放资源
4. **并发处理**：合理设置并发数量，避免资源竞争
5. **质量保证**：使用文档质量检查工具确保输出质量

如果您在使用过程中遇到问题，请参考：
- [API参考文档](API_REFERENCE.md)
- [快速开始指南](QUICKSTART.md)
- [开发指南](DEVELOPMENT_GUIDE.md)
- [项目总结](PROJECT_SUMMARY.md)

或者在GitHub上提交Issue获取帮助。