package examples

import (
    "fmt"
    "log"

    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
    "github.com/tanqiangyes/go-word/pkg/writer"
)

func DemoAdvancedFormatting() {
    fmt.Println("=== Go Word 高级格式化示例 ===")

    // 示例1：复杂表格支持
    demoComplexTable()

    // 示例2：页眉页脚处理
    demoHeaderFooter()

    // 示例3：分页和分节支持
    demoPageAndSection()

    fmt.Println("所有高级格式化示例完成！")
}

// demoComplexTable 演示复杂表格功能
func demoComplexTable() {
    fmt.Println("1. 复杂表格示例")
    fmt.Println("----------------")

    // 创建文档
    w := writer.NewDocumentWriter()
    w.CreateNewDocument()

    // 添加标题
    w.AddParagraph("复杂表格示例", "Heading1")

    // 创建高级格式化器
    doc, err := word.Open("temp_doc.docx")
    if err != nil {
        log.Printf("打开文档失败: %v", err)
        return
    }
    defer doc.Close()

    formatter := word.NewAdvancedFormatter(doc)

    // 创建复杂表格
    table := formatter.CreateComplexTable(4, 3)

    // 设置表格标题
    table.Properties.Caption = "销售数据表"
    table.Properties.Description = "包含销售数据的复杂表格"

    // 设置表头
    table.Rows[0].Header = true
    table.Rows[0].Cells[0].Content.Text = "产品名称"
    table.Rows[0].Cells[1].Content.Text = "销售数量"
    table.Rows[0].Cells[2].Content.Text = "销售金额"

    // 设置数据行
    table.Rows[1].Cells[0].Content.Text = "产品A"
    table.Rows[1].Cells[1].Content.Text = "100"
    table.Rows[1].Cells[2].Content.Text = "10000"

    table.Rows[2].Cells[0].Content.Text = "产品B"
    table.Rows[2].Cells[1].Content.Text = "150"
    table.Rows[2].Cells[2].Content.Text = "15000"

    table.Rows[3].Cells[0].Content.Text = "产品C"
    table.Rows[3].Cells[1].Content.Text = "200"
    table.Rows[3].Cells[2].Content.Text = "20000"

    // 合并单元格（总计行）
    formatter.MergeCells(table, "A4", "C4")
    table.Rows[3].Cells[0].Content.Text = "总计"

    // 设置单元格边框
    borders := word.CellBorders{
        Top:    word.BorderSide{Style: "double", Size: 2, Color: "000000"},
        Bottom: word.BorderSide{Style: "single", Size: 1, Color: "000000"},
        Left:   word.BorderSide{Style: "single", Size: 1, Color: "000000"},
        Right:  word.BorderSide{Style: "single", Size: 1, Color: "000000"},
    }

    formatter.SetCellBorders(table, "A1", borders)
    formatter.SetCellBorders(table, "B1", borders)
    formatter.SetCellBorders(table, "C1", borders)

    // 设置单元格底纹
    shading := word.CellShading{
        Fill:  "solid",
        Color: "CCCCCC",
        Val:   "clear",
    }

    formatter.SetCellShading(table, "A1", shading)
    formatter.SetCellShading(table, "B1", shading)
    formatter.SetCellShading(table, "C1", shading)

    // 添加表格到文档
    if err := formatter.AddComplexTable(table); err != nil {
        log.Printf("添加复杂表格失败: %v", err)
        return
    }

    fmt.Println("✅ 复杂表格创建成功")
    fmt.Printf("表格ID: %s\n", table.ID)
    fmt.Printf("行数: %d, 列数: %d\n", len(table.Rows), len(table.Columns))
    fmt.Printf("标题: %s\n", table.Properties.Caption)
    fmt.Println()
}

// demoHeaderFooter 演示页眉页脚功能
func demoHeaderFooter() {
    fmt.Println("2. 页眉页脚示例")
    fmt.Println("----------------")

    // 创建文档
    w := writer.NewDocumentWriter()
    w.CreateNewDocument()

    // 添加内容
    w.AddParagraph("页眉页脚示例文档", "Heading1")
    w.AddParagraph("这是文档的主要内容。", "Normal")
    w.AddParagraph("文档包含页眉和页脚。", "Normal")

    // 创建高级格式化器
    doc, err := word.Open("temp_doc.docx")
    if err != nil {
        log.Printf("打开文档失败: %v", err)
        return
    }
    defer doc.Close()

    formatter := word.NewAdvancedFormatter(doc)

    // 创建页眉
    header := formatter.CreateHeader(word.HeaderType)
    header.Content[0].Text = "公司名称 - 页眉"
    header.Content[0].Runs[0].Text = "公司名称 - 页眉"
    header.Content[0].Runs[0].Bold = true
    header.Content[0].Runs[0].FontSize = 12

    // 创建页脚
    footer := formatter.CreateFooter(word.FooterType)
    footer.Content[0].Text = "页码: 第 1 页"
    footer.Content[0].Runs[0].Text = "页码: 第 1 页"
    footer.Content[0].Runs[0].FontSize = 10

    // 创建首页页眉
    firstHeader := formatter.CreateHeader(word.FirstHeaderType)
    firstHeader.Content[0].Text = "首页页眉 - 特殊格式"
    firstHeader.Content[0].Runs[0].Text = "首页页眉 - 特殊格式"
    firstHeader.Content[0].Runs[0].Italic = true
    firstHeader.Content[0].Runs[0].FontSize = 14

    // 创建首页页脚
    firstFooter := formatter.CreateFooter(word.FirstFooterType)
    firstFooter.Content[0].Text = "首页页脚 - 文档信息"
    firstFooter.Content[0].Runs[0].Text = "首页页脚 - 文档信息"
    firstFooter.Content[0].Runs[0].FontSize = 10

    // 添加页眉页脚到文档
    if err := formatter.AddHeader(header); err != nil {
        log.Printf("添加页眉失败: %v", err)
        return
    }

    if err := formatter.AddFooter(footer); err != nil {
        log.Printf("添加页脚失败: %v", err)
        return
    }

    if err := formatter.AddHeader(firstHeader); err != nil {
        log.Printf("添加首页页眉失败: %v", err)
        return
    }

    if err := formatter.AddFooter(firstFooter); err != nil {
        log.Printf("添加首页页脚失败: %v", err)
        return
    }

    fmt.Println("✅ 页眉页脚创建成功")
    fmt.Printf("页眉类型: %v\n", header.Type)
    fmt.Printf("页脚类型: %v\n", footer.Type)
    fmt.Printf("首页页眉类型: %v\n", firstHeader.Type)
    fmt.Printf("首页页脚类型: %v\n", firstFooter.Type)
    fmt.Println()
}

// demoPageAndSection 演示分页和分节功能
func demoPageAndSection() {
    fmt.Println("3. 分页和分节示例")
    fmt.Println("------------------")

    // 创建文档
    w := writer.NewDocumentWriter()
    w.CreateNewDocument()

    // 添加内容
    w.AddParagraph("分页和分节示例", "Heading1")
    w.AddParagraph("这是第一页的内容。", "Normal")

    // 创建高级格式化器
    doc, err := word.Open("temp_doc.docx")
    if err != nil {
        log.Printf("打开文档失败: %v", err)
        return
    }
    defer doc.Close()

    formatter := word.NewAdvancedFormatter(doc)

    // 添加分页符
    if err := formatter.AddPageBreak(); err != nil {
        log.Printf("添加分页符失败: %v", err)
        return
    }

    // 创建新节
    section := formatter.CreateSection()
    section.Content = []word.Paragraph{
        {
            Text:  "第二节内容",
            Style: "Heading2",
            Runs: []word.Run{
                {
                    Text:     "第二节内容",
                    FontSize: 16,
                    FontName: "Arial",
                    Bold:     true,
                },
            },
        },
        {
            Text:  "这是第二节的内容，具有不同的页面设置。",
            Style: "Normal",
            Runs: []word.Run{
                {
                    Text:     "这是第二节的内容，具有不同的页面设置。",
                    FontSize: 12,
                    FontName: "Arial",
                },
            },
        },
    }

    // 设置节属性
    section.Properties.PageSize.Width = 612  // 8.5 inches
    section.Properties.PageSize.Height = 792 // 11 inches
    section.Properties.PageSize.Orientation = "portrait"

    section.Properties.PageMargins.Top = 72
    section.Properties.PageMargins.Bottom = 72
    section.Properties.PageMargins.Left = 72
    section.Properties.PageMargins.Right = 72

    section.Properties.PageNumbering.Start = 1
    section.Properties.PageNumbering.Format = "decimal"

    // 添加节到文档
    if err := formatter.AddSection(section); err != nil {
        log.Printf("添加节失败: %v", err)
        return
    }

    // 添加更多分页符
    if err := formatter.AddPageBreak(); err != nil {
        log.Printf("添加分页符失败: %v", err)
        return
    }

    // 添加第三页内容
    w.AddParagraph("第三页内容", "Heading2")
    w.AddParagraph("这是第三页的内容。", "Normal")

    fmt.Println("✅ 分页和分节创建成功")
    fmt.Printf("节ID: %s\n", section.ID)
    fmt.Printf("页面大小: %.0f x %.0f\n", section.Properties.PageSize.Width, section.Properties.PageSize.Height)
    fmt.Printf("页面方向: %s\n", section.Properties.PageSize.Orientation)
    fmt.Printf("页码格式: %s\n", section.Properties.PageNumbering.Format)
    fmt.Println()
}
