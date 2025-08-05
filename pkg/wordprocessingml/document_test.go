package wordprocessingml

import (
	"testing"

	"github.com/tanqiangyes/go-word/pkg/types"
)

func TestDocumentCreation(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	
	if doc == nil {
		t.Fatal("Expected document to be created")
	}
}

func TestDocumentSetMainPart(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	
	doc.SetMainPart(mainPart)
	
	if doc.GetMainPart() != mainPart {
		t.Error("Expected main part to be set correctly")
	}
}

func TestDocumentGetMainPart(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	
	doc.SetMainPart(mainPart)
	retrievedPart := doc.GetMainPart()
	
	if retrievedPart != mainPart {
		t.Error("Expected retrieved main part to match set main part")
	}
}

func TestMainDocumentPartContent(t *testing.T) {
	content := &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{Text: "Test paragraph"},
		},
		Tables: []types.Table{
			{Rows: []types.TableRow{}},
		},
		Text: "Test content",
	}
	
	mainPart := &MainDocumentPart{
		Content: content,
	}
	
	if mainPart.Content != content {
		t.Error("Expected content to be set correctly")
	}
	
	if len(mainPart.Content.Paragraphs) != 1 {
		t.Errorf("Expected 1 paragraph, got %d", len(mainPart.Content.Paragraphs))
	}
	
	if len(mainPart.Content.Tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(mainPart.Content.Tables))
	}
	
	if mainPart.Content.Text != "Test content" {
		t.Errorf("Expected text 'Test content', got '%s'", mainPart.Content.Text)
	}
}

func TestDocumentOpen(t *testing.T) {
	// 测试打开不存在的文档
	_, err := Open("nonexistent.docx")
	if err == nil {
		t.Error("Expected error when opening nonexistent document")
	}
}

func TestDocumentGetContainer(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	container := doc.GetContainer()
	
	// 新创建的文档可能没有容器，这是正常的
	if container != nil {
		t.Log("Container exists for new document")
	} else {
		t.Log("Container is nil for new document (this is normal)")
	}
}

func TestDocumentWithContent(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{
				{Text: "First paragraph"},
				{Text: "Second paragraph"},
			},
			Tables: []types.Table{
				{
					Rows: []types.TableRow{
						{
							Cells: []types.TableCell{
								{Text: "Cell 1"},
								{Text: "Cell 2"},
							},
						},
					},
					Columns: 2,
				},
			},
			Text: "Document text",
		},
	}
	doc.SetMainPart(mainPart)
	
	// 验证内容
	retrievedPart := doc.GetMainPart()
	if len(retrievedPart.Content.Paragraphs) != 2 {
		t.Errorf("Expected 2 paragraphs, got %d", len(retrievedPart.Content.Paragraphs))
	}
	
	if len(retrievedPart.Content.Tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(retrievedPart.Content.Tables))
	}
	
	if retrievedPart.Content.Text != "Document text" {
		t.Errorf("Expected text 'Document text', got '%s'", retrievedPart.Content.Text)
	}
}

func TestDocumentNilMainPart(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	
	// 测试未设置mainPart的情况
	mainPart := doc.GetMainPart()
	if mainPart != nil {
		t.Error("Expected nil main part for new document")
	}
}

func TestMainDocumentPartNilContent(t *testing.T) {
	mainPart := &MainDocumentPart{}
	
	// 测试未设置content的情况
	if mainPart.Content != nil {
		t.Error("Expected nil content for new main part")
	}
}

func TestDocumentGetText(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{
				{Text: "First paragraph"},
				{Text: "Second paragraph"},
			},
			Text: "Document text",
		},
	}
	doc.SetMainPart(mainPart)
	
	text, err := doc.GetText()
	if err != nil {
		t.Fatalf("Failed to get text: %v", err)
	}
	
	if text == "" {
		t.Error("Expected text to not be empty")
	}
}

func TestDocumentGetTextWithNilContent(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	
	// 测试nil mainPart的情况
	_, err := doc.GetText()
	if err == nil {
		t.Error("Expected error when mainPart is nil")
	}
	
	// 测试nil Content的情况
	doc.SetMainPart(&MainDocumentPart{})
	_, err = doc.GetText()
	if err == nil {
		t.Error("Expected error when Content is nil")
	}
}

func TestDocumentGetParagraphs(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{
				{Text: "First paragraph"},
				{Text: "Second paragraph"},
			},
		},
	}
	doc.SetMainPart(mainPart)
	
	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		t.Fatalf("Failed to get paragraphs: %v", err)
	}
	
	if len(paragraphs) != 2 {
		t.Errorf("Expected 2 paragraphs, got %d", len(paragraphs))
	}
}

func TestDocumentGetParagraphsWithNilContent(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	
	// 测试nil mainPart的情况
	_, err := doc.GetParagraphs()
	if err == nil {
		t.Error("Expected error when mainPart is nil")
	}
	
	// 测试nil Content的情况
	doc.SetMainPart(&MainDocumentPart{})
	_, err = doc.GetParagraphs()
	if err == nil {
		t.Error("Expected error when Content is nil")
	}
}

func TestDocumentGetTables(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Tables: []types.Table{
				{Rows: []types.TableRow{}},
			},
		},
	}
	doc.SetMainPart(mainPart)
	
	tables, err := doc.GetTables()
	if err != nil {
		t.Fatalf("Failed to get tables: %v", err)
	}
	
	if len(tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(tables))
	}
}

func TestDocumentGetTablesWithNilContent(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	
	// 测试nil mainPart的情况
	_, err := doc.GetTables()
	if err == nil {
		t.Error("Expected error when mainPart is nil")
	}
	
	// 测试nil Content的情况
	doc.SetMainPart(&MainDocumentPart{})
	_, err = doc.GetTables()
	if err == nil {
		t.Error("Expected error when Content is nil")
	}
}

func TestDocumentGetDocumentParts(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	parts := doc.GetDocumentParts()
	
	// 新创建的文档应该有documentParts
	if parts == nil {
		t.Error("Expected document parts to be created")
	}
}

func TestDocumentGetPartsSummary(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	summary := doc.GetPartsSummary()
	
	if summary == "" {
		t.Error("Expected parts summary to not be empty")
	}
}

func TestDocumentGetPartsSummaryWithNilDocumentParts(t *testing.T) {
	doc := &Document{
		documentParts: nil,
	}
	summary := doc.GetPartsSummary()
	
	if summary != "文档部分未加载" {
		t.Errorf("Expected '文档部分未加载', got '%s'", summary)
	}
}

func TestDocumentClose(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	
	// 测试关闭文档
	err := doc.Close()
	if err != nil {
		t.Errorf("Failed to close document: %v", err)
	}
}

func TestDocumentCloseWithNilContainer(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
		container:     nil,
	}
	
	// 测试关闭没有容器的文档
	err := doc.Close()
	if err != nil {
		t.Errorf("Failed to close document with nil container: %v", err)
	}
}

func TestParseDocumentContent(t *testing.T) {
	// 测试解析有效的XML内容
	xmlContent := []byte(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Test paragraph</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`)
	
	content, err := parseDocumentContent(xmlContent)
	if err != nil {
		t.Fatalf("Failed to parse document content: %v", err)
	}
	
	if content == nil {
		t.Fatal("Expected content to be parsed")
	}
}

func TestParseDocumentContentWithInvalidXML(t *testing.T) {
	// 测试解析无效的XML内容
	invalidXML := []byte(`<invalid xml content`)
	
	_, err := parseDocumentContent(invalidXML)
	if err == nil {
		t.Error("Expected error when parsing invalid XML")
	}
}

func TestLoadMainDocumentPart(t *testing.T) {
	// 这个测试需要模拟OPC容器
	// 由于loadMainDocumentPart是私有方法，我们通过Open方法来测试
	// 这里我们测试Open方法失败的情况
	_, err := Open("nonexistent.docx")
	if err == nil {
		t.Error("Expected error when opening nonexistent document")
	}
}

// DocumentParts 测试
func TestNewDocumentParts(t *testing.T) {
	parts := NewDocumentParts()
	
	if parts == nil {
		t.Fatal("Expected DocumentParts to be created")
	}
	
	// 验证初始化
	if len(parts.HeaderParts) != 0 {
		t.Errorf("Expected 0 header parts, got %d", len(parts.HeaderParts))
	}
	
	if len(parts.FooterParts) != 0 {
		t.Errorf("Expected 0 footer parts, got %d", len(parts.FooterParts))
	}
	
	if len(parts.CommentParts) != 0 {
		t.Errorf("Expected 0 comment parts, got %d", len(parts.CommentParts))
	}
	
	if len(parts.FootnoteParts) != 0 {
		t.Errorf("Expected 0 footnote parts, got %d", len(parts.FootnoteParts))
	}
	
	if len(parts.EndnoteParts) != 0 {
		t.Errorf("Expected 0 endnote parts, got %d", len(parts.EndnoteParts))
	}
}

func TestDocumentPartsAddHeaderPart(t *testing.T) {
	parts := NewDocumentParts()
	
	header := HeaderPart{
		ID:   "header1",
		Type: HeaderType, // 使用正确的常量
		Content: []types.Paragraph{
			{Text: "Header content"},
		},
	}
	
	parts.AddHeaderPart(header)
	
	if len(parts.HeaderParts) != 1 {
		t.Errorf("Expected 1 header part, got %d", len(parts.HeaderParts))
	}
	
	if parts.HeaderParts[0].ID != "header1" {
		t.Errorf("Expected header ID 'header1', got '%s'", parts.HeaderParts[0].ID)
	}
}

func TestDocumentPartsAddFooterPart(t *testing.T) {
	parts := NewDocumentParts()
	
	footer := FooterPart{
		ID:   "footer1",
		Type: FooterType, // 使用正确的常量
		Content: []types.Paragraph{
			{Text: "Footer content"},
		},
	}
	
	parts.AddFooterPart(footer)
	
	if len(parts.FooterParts) != 1 {
		t.Errorf("Expected 1 footer part, got %d", len(parts.FooterParts))
	}
	
	if parts.FooterParts[0].ID != "footer1" {
		t.Errorf("Expected footer ID 'footer1', got '%s'", parts.FooterParts[0].ID)
	}
}

func TestDocumentPartsAddCommentPart(t *testing.T) {
	parts := NewDocumentParts()
	
	comment := CommentPart{
		ID: "comment1",
		Content: []Comment{
			{
				ID:     "comment1",
				Author: "Test Author",
				Text:   "Test comment",
			},
		},
	}
	
	parts.AddCommentPart(comment)
	
	if len(parts.CommentParts) != 1 {
		t.Errorf("Expected 1 comment part, got %d", len(parts.CommentParts))
	}
	
	if parts.CommentParts[0].ID != "comment1" {
		t.Errorf("Expected comment ID 'comment1', got '%s'", parts.CommentParts[0].ID)
	}
}

func TestDocumentPartsAddFootnotePart(t *testing.T) {
	parts := NewDocumentParts()
	
	footnote := FootnotePart{
		ID: "footnote1",
		Content: []Footnote{
			{
				ID:     "footnote1",
				Text:   "Test footnote",
				Number: 1,
			},
		},
	}
	
	parts.AddFootnotePart(footnote)
	
	if len(parts.FootnoteParts) != 1 {
		t.Errorf("Expected 1 footnote part, got %d", len(parts.FootnoteParts))
	}
	
	if parts.FootnoteParts[0].ID != "footnote1" {
		t.Errorf("Expected footnote ID 'footnote1', got '%s'", parts.FootnoteParts[0].ID)
	}
}

func TestDocumentPartsAddEndnotePart(t *testing.T) {
	parts := NewDocumentParts()
	
	endnote := EndnotePart{
		ID: "endnote1",
		Content: []Endnote{
			{
				ID:     "endnote1",
				Text:   "Test endnote",
				Number: 1,
			},
		},
	}
	
	parts.AddEndnotePart(endnote)
	
	if len(parts.EndnoteParts) != 1 {
		t.Errorf("Expected 1 endnote part, got %d", len(parts.EndnoteParts))
	}
	
	if parts.EndnoteParts[0].ID != "endnote1" {
		t.Errorf("Expected endnote ID 'endnote1', got '%s'", parts.EndnoteParts[0].ID)
	}
}

func TestDocumentPartsGetHeaderPart(t *testing.T) {
	parts := NewDocumentParts()
	
	header := HeaderPart{
		ID:   "header1",
		Type: HeaderType, // 使用正确的常量
	}
	
	parts.AddHeaderPart(header)
	
	// 测试获取存在的页眉
	retrievedHeader := parts.GetHeaderPart("header1")
	if retrievedHeader == nil {
		t.Error("Expected header to be found")
	}
	
	if retrievedHeader.ID != "header1" {
		t.Errorf("Expected header ID 'header1', got '%s'", retrievedHeader.ID)
	}
	
	// 测试获取不存在的页眉
	nonexistentHeader := parts.GetHeaderPart("nonexistent")
	if nonexistentHeader != nil {
		t.Error("Expected nil for nonexistent header")
	}
}

func TestDocumentPartsGetFooterPart(t *testing.T) {
	parts := NewDocumentParts()
	
	footer := FooterPart{
		ID:   "footer1",
		Type: FooterType, // 使用正确的常量
	}
	
	parts.AddFooterPart(footer)
	
	// 测试获取存在的页脚
	retrievedFooter := parts.GetFooterPart("footer1")
	if retrievedFooter == nil {
		t.Error("Expected footer to be found")
	}
	
	if retrievedFooter.ID != "footer1" {
		t.Errorf("Expected footer ID 'footer1', got '%s'", retrievedFooter.ID)
	}
	
	// 测试获取不存在的页脚
	nonexistentFooter := parts.GetFooterPart("nonexistent")
	if nonexistentFooter != nil {
		t.Error("Expected nil for nonexistent footer")
	}
}

func TestDocumentPartsGetCommentPart(t *testing.T) {
	parts := NewDocumentParts()
	
	comment := CommentPart{
		ID: "comment1",
		Content: []Comment{
			{
				ID:     "comment1",
				Author: "Test Author",
				Text:   "Test comment",
			},
		},
	}
	
	parts.AddCommentPart(comment)
	
	// 测试获取存在的注释
	retrievedComment := parts.GetCommentPart("comment1")
	if retrievedComment == nil {
		t.Error("Expected comment to be found")
	}
	
	if retrievedComment.ID != "comment1" {
		t.Errorf("Expected comment ID 'comment1', got '%s'", retrievedComment.ID)
	}
	
	// 测试获取不存在的注释
	nonexistentComment := parts.GetCommentPart("nonexistent")
	if nonexistentComment != nil {
		t.Error("Expected nil for nonexistent comment")
	}
}

func TestDocumentPartsGetFootnotePart(t *testing.T) {
	parts := NewDocumentParts()
	
	footnote := FootnotePart{
		ID: "footnote1",
		Content: []Footnote{
			{
				ID:     "footnote1",
				Text:   "Test footnote",
				Number: 1,
			},
		},
	}
	
	parts.AddFootnotePart(footnote)
	
	// 测试获取存在的脚注
	retrievedFootnote := parts.GetFootnotePart("footnote1")
	if retrievedFootnote == nil {
		t.Error("Expected footnote to be found")
	}
	
	if retrievedFootnote.ID != "footnote1" {
		t.Errorf("Expected footnote ID 'footnote1', got '%s'", retrievedFootnote.ID)
	}
	
	// 测试获取不存在的脚注
	nonexistentFootnote := parts.GetFootnotePart("nonexistent")
	if nonexistentFootnote != nil {
		t.Error("Expected nil for nonexistent footnote")
	}
}

func TestDocumentPartsGetEndnotePart(t *testing.T) {
	parts := NewDocumentParts()
	
	endnote := EndnotePart{
		ID: "endnote1",
		Content: []Endnote{
			{
				ID:     "endnote1",
				Text:   "Test endnote",
				Number: 1,
			},
		},
	}
	
	parts.AddEndnotePart(endnote)
	
	// 测试获取存在的尾注
	retrievedEndnote := parts.GetEndnotePart("endnote1")
	if retrievedEndnote == nil {
		t.Error("Expected endnote to be found")
	}
	
	if retrievedEndnote.ID != "endnote1" {
		t.Errorf("Expected endnote ID 'endnote1', got '%s'", retrievedEndnote.ID)
	}
	
	// 测试获取不存在的尾注
	nonexistentEndnote := parts.GetEndnotePart("nonexistent")
	if nonexistentEndnote != nil {
		t.Error("Expected nil for nonexistent endnote")
	}
}

func TestDocumentPartsGetPartsSummary(t *testing.T) {
	parts := NewDocumentParts()
	
	// 添加一些部分
	parts.AddHeaderPart(HeaderPart{ID: "header1"})
	parts.AddFooterPart(FooterPart{ID: "footer1"})
	parts.AddCommentPart(CommentPart{ID: "comment1"})
	parts.AddFootnotePart(FootnotePart{ID: "footnote1"})
	parts.AddEndnotePart(EndnotePart{ID: "endnote1"})
	
	summary := parts.GetPartsSummary()
	
	if summary == "" {
		t.Error("Expected summary to not be empty")
	}
	
	// 验证摘要包含预期的内容
	expectedContent := []string{
		"文档部分摘要:",
		"页眉部分: 1",
		"页脚部分: 1",
		"注释部分: 1",
		"脚注部分: 1",
		"尾注部分: 1",
	}
	
	for _, expected := range expectedContent {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

func TestDocumentPartsGetPartsSummaryWithNilParts(t *testing.T) {
	parts := &DocumentParts{}
	
	summary := parts.GetPartsSummary()
	
	if summary == "" {
		t.Error("Expected summary to not be empty")
	}
	
	// 验证摘要包含预期的内容
	expectedContent := []string{
		"文档部分摘要:",
		"页眉部分: 0",
		"页脚部分: 0",
		"注释部分: 0",
		"脚注部分: 0",
		"尾注部分: 0",
	}
	
	for _, expected := range expectedContent {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

// 辅助函数
func contains(s, substr string) bool {
	return len(s) >= len(substr) && (s == substr || len(s) > len(substr) && (s[:len(substr)] == substr || s[len(s)-len(substr):] == substr || containsSubstring(s, substr)))
}

func containsSubstring(s, substr string) bool {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}

// 高级格式化测试
func TestNewAdvancedFormatter(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	
	formatter := NewAdvancedFormatter(doc)
	
	if formatter == nil {
		t.Fatal("Expected AdvancedFormatter to be created")
	}
	
	if formatter.Document != doc {
		t.Error("Expected formatter to reference the document")
	}
}

func TestAdvancedFormatterCreateComplexTable(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	// 设置mainPart以避免panic
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	
	formatter := NewAdvancedFormatter(doc)
	
	// 创建一个3x3的复杂表格
	table := formatter.CreateComplexTable(3, 3)
	
	if table == nil {
		t.Fatal("Expected table to be created")
	}
	
	if len(table.Rows) != 3 {
		t.Errorf("Expected 3 rows, got %d", len(table.Rows))
	}
	
	if len(table.Columns) != 3 {
		t.Errorf("Expected 3 columns, got %d", len(table.Columns))
	}
	
	// 验证行
	for i, row := range table.Rows {
		if row.Index != i+1 {
			t.Errorf("Expected row index %d, got %d", i+1, row.Index)
		}
		
		if len(row.Cells) != 3 {
			t.Errorf("Expected 3 cells in row %d, got %d", i+1, len(row.Cells))
		}
	}
	
	// 验证列
	for i, col := range table.Columns {
		if col.Index != i+1 {
			t.Errorf("Expected column index %d, got %d", i+1, col.Index)
		}
		
		if col.Width != 20 {
			t.Errorf("Expected column width 20, got %f", col.Width)
		}
	}
}

func TestAdvancedFormatterAddComplexTable(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	
	formatter := NewAdvancedFormatter(doc)
	table := formatter.CreateComplexTable(2, 2)
	
	err := formatter.AddComplexTable(table)
	if err != nil {
		t.Fatalf("Failed to add complex table: %v", err)
	}
	
	// 验证表格已添加到文档
	retrievedPart := doc.GetMainPart()
	if len(retrievedPart.Content.Tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(retrievedPart.Content.Tables))
	}
}

func TestAdvancedFormatterAddComplexTableWithNilDocument(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	formatter := NewAdvancedFormatter(doc)
	
	// 创建一个简单的表格而不调用CreateComplexTable（因为它需要mainPart）
	table := &ComplexTable{
		ID: "test_table",
		Rows: []ComplexTableRow{
			{
				Index: 1,
				Cells: []ComplexTableCell{
					{
						Reference: "A1",
						Content: CellContent{
							Text: "Test cell",
						},
					},
				},
			},
		},
		Columns: []ComplexTableColumn{
			{
				Index: 1,
				Width: 20,
			},
		},
	}
	
	err := formatter.AddComplexTable(table)
	if err == nil {
		t.Error("Expected error when document content is nil")
	}
}

func TestAdvancedFormatterCreateHeader(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	formatter := NewAdvancedFormatter(doc)
	
	header := formatter.CreateHeader(HeaderType)
	
	if header == nil {
		t.Fatal("Expected header to be created")
	}
	
	if header.Type != HeaderType {
		t.Errorf("Expected header type HeaderType, got %v", header.Type)
	}
	
	if len(header.Content) == 0 {
		t.Error("Expected header to have content")
	}
	
	if header.Properties.DifferentFirst != false {
		t.Error("Expected DifferentFirst to be false")
	}
	
	if header.Properties.DifferentOddEven != false {
		t.Error("Expected DifferentOddEven to be false")
	}
	
	if header.Properties.AlignWithMargins != true {
		t.Error("Expected AlignWithMargins to be true")
	}
	
	if header.Properties.ScaleWithDoc != true {
		t.Error("Expected ScaleWithDoc to be true")
	}
}

func TestAdvancedFormatterCreateFooter(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	formatter := NewAdvancedFormatter(doc)
	
	footer := formatter.CreateFooter(FooterType)
	
	if footer == nil {
		t.Fatal("Expected footer to be created")
	}
	
	if footer.Type != FooterType {
		t.Errorf("Expected footer type FooterType, got %v", footer.Type)
	}
	
	if len(footer.Content) == 0 {
		t.Error("Expected footer to have content")
	}
	
	if footer.Properties.DifferentFirst != false {
		t.Error("Expected DifferentFirst to be false")
	}
	
	if footer.Properties.DifferentOddEven != false {
		t.Error("Expected DifferentOddEven to be false")
	}
	
	if footer.Properties.AlignWithMargins != true {
		t.Error("Expected AlignWithMargins to be true")
	}
	
	if footer.Properties.ScaleWithDoc != true {
		t.Error("Expected ScaleWithDoc to be true")
	}
}

func TestAdvancedFormatterAddHeader(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	
	formatter := NewAdvancedFormatter(doc)
	header := formatter.CreateHeader(HeaderType)
	
	err := formatter.AddHeader(header)
	if err != nil {
		t.Fatalf("Failed to add header: %v", err)
	}
	
	// 验证页眉内容已添加到文档
	retrievedPart := doc.GetMainPart()
	if len(retrievedPart.Content.Paragraphs) == 0 {
		t.Error("Expected header paragraphs to be added to document")
	}
}

func TestAdvancedFormatterAddHeaderWithNilDocument(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	formatter := NewAdvancedFormatter(doc)
	header := formatter.CreateHeader(HeaderType)
	
	err := formatter.AddHeader(header)
	if err == nil {
		t.Error("Expected error when document content is nil")
	}
}

func TestAdvancedFormatterAddFooter(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	
	formatter := NewAdvancedFormatter(doc)
	footer := formatter.CreateFooter(FooterType)
	
	err := formatter.AddFooter(footer)
	if err != nil {
		t.Fatalf("Failed to add footer: %v", err)
	}
	
	// 验证页脚内容已添加到文档
	retrievedPart := doc.GetMainPart()
	if len(retrievedPart.Content.Paragraphs) == 0 {
		t.Error("Expected footer paragraphs to be added to document")
	}
}

func TestAdvancedFormatterAddFooterWithNilDocument(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	formatter := NewAdvancedFormatter(doc)
	footer := formatter.CreateFooter(FooterType)
	
	err := formatter.AddFooter(footer)
	if err == nil {
		t.Error("Expected error when document content is nil")
	}
}

func TestAdvancedFormatterCreateSection(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	// 设置mainPart以避免panic
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	
	formatter := NewAdvancedFormatter(doc)
	
	section := formatter.CreateSection()
	
	if section == nil {
		t.Fatal("Expected section to be created")
	}
	
	if section.ID == "" {
		t.Error("Expected section to have an ID")
	}
	
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
	
	if section.Properties.PageMargins.Bottom != 72 {
		t.Errorf("Expected bottom margin 72, got %f", section.Properties.PageMargins.Bottom)
	}
	
	if section.Properties.PageMargins.Left != 72 {
		t.Errorf("Expected left margin 72, got %f", section.Properties.PageMargins.Left)
	}
	
	if section.Properties.PageMargins.Right != 72 {
		t.Errorf("Expected right margin 72, got %f", section.Properties.PageMargins.Right)
	}
}

func TestAdvancedFormatterAddSection(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	
	formatter := NewAdvancedFormatter(doc)
	section := formatter.CreateSection()
	
	err := formatter.AddSection(section)
	if err != nil {
		t.Fatalf("Failed to add section: %v", err)
	}
	
	// 验证节内容已添加到文档
	retrievedPart := doc.GetMainPart()
	if len(retrievedPart.Content.Paragraphs) == 0 {
		t.Error("Expected section paragraphs to be added to document")
	}
}

func TestAdvancedFormatterAddSectionWithNilDocument(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	formatter := NewAdvancedFormatter(doc)
	
	// 创建一个简单的section而不调用CreateSection（因为它需要mainPart）
	section := &Section{
		ID: "test_section",
		Properties: SectionProperties{
			PageSize: PageSize{
				Width:       612,
				Height:      792,
				Orientation: "portrait",
			},
			PageMargins: PageMargins{
				Top:    72,
				Bottom: 72,
				Left:   72,
				Right:  72,
			},
		},
		Content: []types.Paragraph{
			{
				Text: "Test section content",
			},
		},
	}
	
	err := formatter.AddSection(section)
	if err == nil {
		t.Error("Expected error when document content is nil")
	}
}

func TestAdvancedFormatterAddPageBreak(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	
	formatter := NewAdvancedFormatter(doc)
	
	err := formatter.AddPageBreak()
	if err != nil {
		t.Fatalf("Failed to add page break: %v", err)
	}
	
	// 验证分页符已添加到文档
	retrievedPart := doc.GetMainPart()
	if len(retrievedPart.Content.Paragraphs) == 0 {
		t.Error("Expected page break paragraph to be added to document")
	}
}

func TestAdvancedFormatterAddPageBreakWithNilDocument(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	formatter := NewAdvancedFormatter(doc)
	
	err := formatter.AddPageBreak()
	if err == nil {
		t.Error("Expected error when document content is nil")
	}
}

func TestAdvancedFormatterMergeCells(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	// 设置mainPart以避免panic
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	
	formatter := NewAdvancedFormatter(doc)
	table := formatter.CreateComplexTable(3, 3)
	
	// 合并A1到B2的单元格
	err := formatter.MergeCells(table, "A1", "B2")
	if err != nil {
		t.Fatalf("Failed to merge cells: %v", err)
	}
	
	// 验证起始单元格（A1）没有被标记为合并，但设置了合并信息
	startCell := table.Rows[0].Cells[0]
	if startCell.Merged {
		t.Error("Start cell should not be marked as merged")
	}
	
	if startCell.MergeStart != "A1" {
		t.Errorf("Expected merge start 'A1', got '%s'", startCell.MergeStart)
	}
	
	if startCell.MergeEnd != "B2" {
		t.Errorf("Expected merge end 'B2', got '%s'", startCell.MergeEnd)
	}
	
	// 验证其他被合并的单元格
	mergedCell := table.Rows[0].Cells[1] // B1
	if !mergedCell.Merged {
		t.Error("Expected cell B1 to be marked as merged")
	}
	
	if mergedCell.MergeStart != "A1" {
		t.Errorf("Expected merge start 'A1', got '%s'", mergedCell.MergeStart)
	}
	
	if mergedCell.MergeEnd != "B2" {
		t.Errorf("Expected merge end 'B2', got '%s'", mergedCell.MergeEnd)
	}
}

func TestAdvancedFormatterMergeCellsWithInvalidReference(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	// 设置mainPart以避免panic
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	formatter := NewAdvancedFormatter(doc)
	table := formatter.CreateComplexTable(3, 3)
	
	// 测试无效的单元格引用
	err := formatter.MergeCells(table, "1A", "A1")
	if err == nil {
		t.Error("Expected error when using invalid cell reference")
	}
}

func TestAdvancedFormatterSetCellBorders(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	// 设置mainPart以避免panic
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	formatter := NewAdvancedFormatter(doc)
	table := formatter.CreateComplexTable(3, 3)
	
	borders := CellBorders{
		Top: BorderSide{
			Style: "single",
			Size:  1,
			Color: "000000",
		},
		Bottom: BorderSide{
			Style: "single",
			Size:  1,
			Color: "000000",
		},
	}
	
	err := formatter.SetCellBorders(table, "A1", borders)
	if err != nil {
		t.Fatalf("Failed to set cell borders: %v", err)
	}
	
	// 验证边框已设置
	cell := table.Rows[0].Cells[0]
	if cell.Borders.Top.Style != "single" {
		t.Errorf("Expected top border style 'single', got '%s'", cell.Borders.Top.Style)
	}
	
	if cell.Borders.Bottom.Style != "single" {
		t.Errorf("Expected bottom border style 'single', got '%s'", cell.Borders.Bottom.Style)
	}
}

func TestAdvancedFormatterSetCellBordersWithInvalidReference(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	// 设置mainPart以避免panic
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	formatter := NewAdvancedFormatter(doc)
	table := formatter.CreateComplexTable(3, 3)
	
	borders := CellBorders{}
	
	err := formatter.SetCellBorders(table, "Z1", borders)
	if err == nil {
		t.Error("Expected error when using invalid cell reference")
	}
}

func TestAdvancedFormatterSetCellShading(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	// 设置mainPart以避免panic
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	formatter := NewAdvancedFormatter(doc)
	table := formatter.CreateComplexTable(3, 3)
	
	shading := CellShading{
		Fill:  "solid",
		Color: "FFFF00",
		Val:   "yellow",
	}
	
	err := formatter.SetCellShading(table, "A1", shading)
	if err != nil {
		t.Fatalf("Failed to set cell shading: %v", err)
	}
	
	// 验证底纹已设置
	cell := table.Rows[0].Cells[0]
	if cell.Shading.Fill != "solid" {
		t.Errorf("Expected shading fill 'solid', got '%s'", cell.Shading.Fill)
	}
	
	if cell.Shading.Color != "FFFF00" {
		t.Errorf("Expected shading color 'FFFF00', got '%s'", cell.Shading.Color)
	}
	
	if cell.Shading.Val != "yellow" {
		t.Errorf("Expected shading val 'yellow', got '%s'", cell.Shading.Val)
	}
}

func TestAdvancedFormatterSetCellShadingWithInvalidReference(t *testing.T) {
	doc := &Document{
		documentParts: NewDocumentParts(),
	}
	// 设置mainPart以避免panic
	mainPart := &MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	doc.SetMainPart(mainPart)
	formatter := NewAdvancedFormatter(doc)
	table := formatter.CreateComplexTable(3, 3)
	
	shading := CellShading{}
	
	err := formatter.SetCellShading(table, "Z1", shading)
	if err == nil {
		t.Error("Expected error when using invalid cell reference")
	}
}

// 测试parseCellReference函数
func TestParseCellReference(t *testing.T) {
	// 测试有效的单元格引用
	col, row, err := parseCellReference("A1")
	if err != nil {
		t.Fatalf("Failed to parse cell reference 'A1': %v", err)
	}
	
	if col != 1 {
		t.Errorf("Expected column 1, got %d", col)
	}
	
	if row != 1 {
		t.Errorf("Expected row 1, got %d", row)
	}
	
	// 测试B2
	col, row, err = parseCellReference("B2")
	if err != nil {
		t.Fatalf("Failed to parse cell reference 'B2': %v", err)
	}
	
	if col != 2 {
		t.Errorf("Expected column 2, got %d", col)
	}
	
	if row != 2 {
		t.Errorf("Expected row 2, got %d", row)
	}
	
	// 测试Z26
	col, row, err = parseCellReference("Z26")
	if err != nil {
		t.Fatalf("Failed to parse cell reference 'Z26': %v", err)
	}
	
	if col != 26 {
		t.Errorf("Expected column 26, got %d", col)
	}
	
	if row != 26 {
		t.Errorf("Expected row 26, got %d", row)
	}
}

func TestParseCellReferenceWithInvalidInput(t *testing.T) {
	// 测试无效的单元格引用
	_, _, err := parseCellReference("")
	if err == nil {
		t.Error("Expected error for empty cell reference")
	}
	
	_, _, err = parseCellReference("1A")
	if err == nil {
		t.Error("Expected error for invalid cell reference '1A'")
	}
	
	_, _, err = parseCellReference("A")
	if err == nil {
		t.Error("Expected error for invalid cell reference 'A'")
	}
	
	_, _, err = parseCellReference("1")
	if err == nil {
		t.Error("Expected error for invalid cell reference '1'")
	}
	
	_, _, err = parseCellReference("A@1")
	if err == nil {
		t.Error("Expected error for invalid cell reference 'A@1'")
	}
}

// 测试复杂表格结构
func TestComplexTableStructure(t *testing.T) {
	table := &ComplexTable{
		ID: "table1",
		Rows: []ComplexTableRow{
			{
				Index: 1,
				Cells: []ComplexTableCell{
					{
						Reference: "A1",
						Content: CellContent{
							Text: "Cell A1",
						},
					},
					{
						Reference: "B1",
						Content: CellContent{
							Text: "Cell B1",
						},
					},
				},
			},
		},
		Columns: []ComplexTableColumn{
			{
				Index: 1,
				Width: 20,
			},
			{
				Index: 2,
				Width: 30,
			},
		},
		Properties: TableProperties{
			Width:     100,
			Alignment: "left",
		},
	}
	
	if table.ID != "table1" {
		t.Errorf("Expected table ID 'table1', got '%s'", table.ID)
	}
	
	if len(table.Rows) != 1 {
		t.Errorf("Expected 1 row, got %d", len(table.Rows))
	}
	
	if len(table.Columns) != 2 {
		t.Errorf("Expected 2 columns, got %d", len(table.Columns))
	}
	
	if table.Properties.Width != 100 {
		t.Errorf("Expected table width 100, got %f", table.Properties.Width)
	}
	
	if table.Properties.Alignment != "left" {
		t.Errorf("Expected table alignment 'left', got '%s'", table.Properties.Alignment)
	}
}

// 测试边框和底纹结构
func TestBorderAndShadingStructures(t *testing.T) {
	border := BorderSide{
		Style:  "single",
		Size:   1,
		Color:  "000000",
		Space:  0,
		Shadow: false,
	}
	
	if border.Style != "single" {
		t.Errorf("Expected border style 'single', got '%s'", border.Style)
	}
	
	if border.Size != 1 {
		t.Errorf("Expected border size 1, got %d", border.Size)
	}
	
	if border.Color != "000000" {
		t.Errorf("Expected border color '000000', got '%s'", border.Color)
	}
	
	shading := CellShading{
		Fill:      "solid",
		Color:     "FFFF00",
		ThemeFill: "accent1",
		ThemeColor: "accent1",
		Val:       "yellow",
	}
	
	if shading.Fill != "solid" {
		t.Errorf("Expected shading fill 'solid', got '%s'", shading.Fill)
	}
	
	if shading.Color != "FFFF00" {
		t.Errorf("Expected shading color 'FFFF00', got '%s'", shading.Color)
	}
	
	if shading.Val != "yellow" {
		t.Errorf("Expected shading val 'yellow', got '%s'", shading.Val)
	}
}

// 测试图片结构
func TestImageStructure(t *testing.T) {
	image := Image{
		ID:          "img1",
		Path:        "/images/test.png",
		Width:       100,
		Height:      200,
		AltText:     "Test Image",
		Title:       "Test Image Title",
		Description: "This is a test image",
	}
	
	if image.ID != "img1" {
		t.Errorf("Expected image ID 'img1', got '%s'", image.ID)
	}
	
	if image.Path != "/images/test.png" {
		t.Errorf("Expected image path '/images/test.png', got '%s'", image.Path)
	}
	
	if image.Width != 100 {
		t.Errorf("Expected image width 100, got %f", image.Width)
	}
	
	if image.Height != 200 {
		t.Errorf("Expected image height 200, got %f", image.Height)
	}
	
	if image.AltText != "Test Image" {
		t.Errorf("Expected image alt text 'Test Image', got '%s'", image.AltText)
	}
	
	if image.Title != "Test Image Title" {
		t.Errorf("Expected image title 'Test Image Title', got '%s'", image.Title)
	}
	
	if image.Description != "This is a test image" {
		t.Errorf("Expected image description 'This is a test image', got '%s'", image.Description)
	}
}

// 测试页面设置结构
func TestPageSettingsStructures(t *testing.T) {
	pageSize := PageSize{
		Width:       11906,
		Height:      16838,
		Orientation: "portrait",
	}
	
	if pageSize.Width != 11906 {
		t.Errorf("Expected page width 11906, got %f", pageSize.Width)
	}
	
	if pageSize.Height != 16838 {
		t.Errorf("Expected page height 16838, got %f", pageSize.Height)
	}
	
	if pageSize.Orientation != "portrait" {
		t.Errorf("Expected page orientation 'portrait', got '%s'", pageSize.Orientation)
	}
	
	pageMargins := PageMargins{
		Top:    1440,
		Bottom: 1440,
		Left:   1800,
		Right:  1800,
		Header: 720,
		Footer: 720,
		Gutter: 0,
	}
	
	if pageMargins.Top != 1440 {
		t.Errorf("Expected top margin 1440, got %f", pageMargins.Top)
	}
	
	if pageMargins.Bottom != 1440 {
		t.Errorf("Expected bottom margin 1440, got %f", pageMargins.Bottom)
	}
	
	if pageMargins.Left != 1800 {
		t.Errorf("Expected left margin 1800, got %f", pageMargins.Left)
	}
	
	if pageMargins.Right != 1800 {
		t.Errorf("Expected right margin 1800, got %f", pageMargins.Right)
	}
} 