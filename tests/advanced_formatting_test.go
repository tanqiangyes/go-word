package tests

import (
	"testing"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/word"
)

func TestNewAdvancedFormatter(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})

	// 创建高级格式化器
	formatter := word.NewAdvancedFormatter(doc)

	if formatter == nil {
		t.Fatal("Expected formatter to be created")
	}

	if formatter.Document != doc {
		t.Error("Expected formatter to have the correct document")
	}
}

func TestCreateComplexTable(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建复杂表格
	table := formatter.CreateComplexTable(3, 4)

	if table == nil {
		t.Fatal("Expected table to be created")
	}

	if len(table.Rows) != 3 {
		t.Errorf("Expected 3 rows, got %d", len(table.Rows))
	}

	if len(table.Columns) != 4 {
		t.Errorf("Expected 4 columns, got %d", len(table.Columns))
	}

	// 验证表格属性
	if table.Properties.Width != 100 {
		t.Errorf("Expected default width 100, got %f", table.Properties.Width)
	}
}

func TestAddComplexTable(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建复杂表格
	table := formatter.CreateComplexTable(2, 3)

	// 添加表格到文档
	err := formatter.AddComplexTable(table)
	if err != nil {
		t.Fatalf("Failed to add complex table: %v", err)
	}

	// 验证表格已添加
	if len(doc.GetMainPart().Content.Tables) == 0 {
		t.Error("Expected table to be added to document")
	}
}

func TestCreateHeader(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	formatter := word.NewAdvancedFormatter(doc)

	// 创建页眉
	header := formatter.CreateHeader(word.HeaderType)

	if header == nil {
		t.Fatal("Expected header to be created")
	}

	if header.Type != word.HeaderType {
		t.Errorf("Expected header type %d, got %d", word.HeaderType, header.Type)
	}
}

func TestCreateFooter(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	formatter := word.NewAdvancedFormatter(doc)

	// 创建页脚
	footer := formatter.CreateFooter(word.FooterType)

	if footer == nil {
		t.Fatal("Expected footer to be created")
	}

	if footer.Type != word.FooterType {
		t.Errorf("Expected footer type %d, got %d", word.FooterType, footer.Type)
	}
}

func TestAddHeader(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建页眉
	header := formatter.CreateHeader(word.HeaderType)

	// 添加页眉到文档
	err := formatter.AddHeader(header)
	if err != nil {
		t.Fatalf("Failed to add header: %v", err)
	}
}

func TestAddFooter(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建页脚
	footer := formatter.CreateFooter(word.FooterType)

	// 添加页脚到文档
	err := formatter.AddFooter(footer)
	if err != nil {
		t.Fatalf("Failed to add footer: %v", err)
	}
}

func TestCreateSection(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建分节
	section := formatter.CreateSection()

	if section == nil {
		t.Fatal("Expected section to be created")
	}

	// 验证分节属性
	if section.Properties.PageSize.Width != 612 {
		t.Errorf("Expected page width 612, got %f", section.Properties.PageSize.Width)
	}

	if section.Properties.PageSize.Height != 792 {
		t.Errorf("Expected page height 792, got %f", section.Properties.PageSize.Height)
	}
}

func TestAddSection(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建分节
	section := formatter.CreateSection()

	// 添加分节到文档
	err := formatter.AddSection(section)
	if err != nil {
		t.Fatalf("Failed to add section: %v", err)
	}
}

func TestAddPageBreak(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 添加分页符
	err := formatter.AddPageBreak()
	if err != nil {
		t.Fatalf("Failed to add page break: %v", err)
	}
}

func TestMergeCells(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建复杂表格
	table := formatter.CreateComplexTable(3, 3)

	// 合并单元格
	err := formatter.MergeCells(table, "A1", "B2")
	if err != nil {
		t.Fatalf("Failed to merge cells: %v", err)
	}

	// 验证合并结果
	// 这里可以添加更多验证逻辑
}

func TestSetCellBorders(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建复杂表格
	table := formatter.CreateComplexTable(2, 2)

	// 设置单元格边框
	borders := word.CellBorders{
		Top: word.BorderSide{
			Style: "single",
			Size:  4,
			Color: "000000",
		},
	}

	err := formatter.SetCellBorders(table, "A1", borders)
	if err != nil {
		t.Fatalf("Failed to set cell borders: %v", err)
	}
}

func TestSetCellShading(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建复杂表格
	table := formatter.CreateComplexTable(2, 2)

	// 设置单元格底纹
	shading := word.CellShading{
		Fill:  "solid",
		Color: "FFFF00",
	}

	err := formatter.SetCellShading(table, "A1", shading)
	if err != nil {
		t.Fatalf("Failed to set cell shading: %v", err)
	}
}

func TestTableProperties(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建复杂表格
	table := formatter.CreateComplexTable(2, 2)

	// 设置表格属性
	table.Properties.Width = 500
	table.Properties.Alignment = "center"
	table.Properties.Caption = "Test Table"

	// 验证表格属性
	if table.Properties.Width != 500 {
		t.Errorf("Expected width 500, got %f", table.Properties.Width)
	}

	if table.Properties.Alignment != "center" {
		t.Errorf("Expected alignment 'center', got '%s'", table.Properties.Alignment)
	}

	if table.Properties.Caption != "Test Table" {
		t.Errorf("Expected caption 'Test Table', got '%s'", table.Properties.Caption)
	}
}

func TestHeaderFooterProperties(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建页眉
	header := formatter.CreateHeader(word.HeaderType)

	// 设置页眉属性
	header.Properties.DifferentFirst = true
	header.Properties.DifferentOddEven = false
	header.Properties.AlignWithMargins = true
	header.Properties.ScaleWithDoc = true

	// 验证页眉属性
	if !header.Properties.DifferentFirst {
		t.Error("Expected DifferentFirst to be true")
	}

	if header.Properties.DifferentOddEven {
		t.Error("Expected DifferentOddEven to be false")
	}

	if !header.Properties.AlignWithMargins {
		t.Error("Expected AlignWithMargins to be true")
	}

	if !header.Properties.ScaleWithDoc {
		t.Error("Expected ScaleWithDoc to be true")
	}
}

func TestSectionProperties(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建分节
	section := formatter.CreateSection()

	// 设置分节属性
	section.Properties.PageSize.Width = 612
	section.Properties.PageSize.Height = 792
	section.Properties.PageSize.Orientation = "portrait"

	section.Properties.PageMargins.Top = 72
	section.Properties.PageMargins.Bottom = 72
	section.Properties.PageMargins.Left = 72
	section.Properties.PageMargins.Right = 72

	// 验证分节属性
	if section.Properties.PageSize.Width != 612 {
		t.Errorf("Expected page width 612, got %f", section.Properties.PageSize.Width)
	}

	if section.Properties.PageSize.Height != 792 {
		t.Errorf("Expected page height 792, got %f", section.Properties.PageSize.Height)
	}

	if section.Properties.PageSize.Orientation != "portrait" {
		t.Errorf("Expected orientation 'portrait', got '%s'", section.Properties.PageSize.Orientation)
	}

	if section.Properties.PageMargins.Top != 72 {
		t.Errorf("Expected top margin 72, got %f", section.Properties.PageMargins.Top)
	}
}

func TestPageNumbering(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建分节
	section := formatter.CreateSection()

	// 设置页码
	section.Properties.PageNumbering.Start = 1
	section.Properties.PageNumbering.Format = "decimal"
	section.Properties.PageNumbering.Restart = "newPage"

	// 验证页码设置
	if section.Properties.PageNumbering.Start != 1 {
		t.Errorf("Expected page numbering start 1, got %d", section.Properties.PageNumbering.Start)
	}

	if section.Properties.PageNumbering.Format != "decimal" {
		t.Errorf("Expected page numbering format 'decimal', got '%s'", section.Properties.PageNumbering.Format)
	}

	if section.Properties.PageNumbering.Restart != "newPage" {
		t.Errorf("Expected page numbering restart 'newPage', got '%s'", section.Properties.PageNumbering.Restart)
	}
}

func TestLineNumbering(t *testing.T) {
	// 创建文档
	doc := &word.Document{}
	doc.SetMainPart(&word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	})
	formatter := word.NewAdvancedFormatter(doc)

	// 创建分节
	section := formatter.CreateSection()

	// 设置行号
	section.Properties.LineNumbering.Start = 1
	section.Properties.LineNumbering.CountBy = 1
	section.Properties.LineNumbering.Distance = 720
	section.Properties.LineNumbering.Restart = "newPage"

	// 验证行号设置
	if section.Properties.LineNumbering.Start != 1 {
		t.Errorf("Expected line numbering start 1, got %d", section.Properties.LineNumbering.Start)
	}

	if section.Properties.LineNumbering.CountBy != 1 {
		t.Errorf("Expected line numbering count by 1, got %d", section.Properties.LineNumbering.CountBy)
	}

	if section.Properties.LineNumbering.Distance != 720 {
		t.Errorf("Expected line numbering distance 720, got %f", section.Properties.LineNumbering.Distance)
	}

	if section.Properties.LineNumbering.Restart != "newPage" {
		t.Errorf("Expected line numbering restart 'newPage', got '%s'", section.Properties.LineNumbering.Restart)
	}
}
