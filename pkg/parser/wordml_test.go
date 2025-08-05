package parser

import (
	"os"
	"testing"

	"github.com/tanqiangyes/go-word/pkg/types"
)

func TestWordMLParserParseDocument(t *testing.T) {
	// 读取测试文档
	content, err := os.ReadFile("../../tests/test_document.xml")
	if err != nil {
		t.Fatalf("Failed to read test document: %v", err)
	}
	
	parser := &WordMLParser{}
	doc, err := parser.ParseWordDocument(content)
	if err != nil {
		t.Fatalf("Failed to parse document: %v", err)
	}
	
	// 验证文档结构
	if doc.Body.Paragraphs == nil {
		t.Error("Expected paragraphs to be parsed")
	}
	
	if len(doc.Body.Paragraphs) != 2 {
		t.Errorf("Expected 2 paragraphs, got %d", len(doc.Body.Paragraphs))
	}
	
	if len(doc.Body.Tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(doc.Body.Tables))
	}
}

func TestWordMLParserParseDocumentWithInvalidXML(t *testing.T) {
	parser := &WordMLParser{}
	
	// 测试无效的XML
	invalidXML := []byte(`<invalid xml content`)
	_, err := parser.ParseWordDocument(invalidXML)
	if err == nil {
		t.Error("Expected error when parsing invalid XML")
	}
}

func TestWordMLParserParseDocumentWithEmptyContent(t *testing.T) {
	parser := &WordMLParser{}
	
	// 测试空内容
	_, err := parser.ParseWordDocument([]byte{})
	if err == nil {
		t.Error("Expected error when parsing empty content")
	}
}

func TestWordMLParserExtractText(t *testing.T) {
	content, err := os.ReadFile("../../tests/test_document.xml")
	if err != nil {
		t.Fatalf("Failed to read test document: %v", err)
	}
	
	parser := &WordMLParser{}
	doc, err := parser.ParseWordDocument(content)
	if err != nil {
		t.Fatalf("Failed to parse document: %v", err)
	}
	
	text := parser.ExtractText(doc)
	
	// 验证提取的文本
	expectedText := "这是一个测试文档\n这是第二段，包含斜体文本。\n"
	if text != expectedText {
		t.Errorf("Expected text '%s', got '%s'", expectedText, text)
	}
}

func TestWordMLParserExtractTextWithEmptyDocument(t *testing.T) {
	parser := &WordMLParser{}
	doc := &WordDocument{
		Body: WordBody{
			Paragraphs: []WordParagraph{},
			Tables:     []WordTable{},
		},
	}
	
	text := parser.ExtractText(doc)
	if text != "" {
		t.Errorf("Expected empty text, got '%s'", text)
	}
}

func TestWordMLParserExtractParagraphs(t *testing.T) {
	content, err := os.ReadFile("../../tests/test_document.xml")
	if err != nil {
		t.Fatalf("Failed to read test document: %v", err)
	}
	
	parser := &WordMLParser{}
	doc, err := parser.ParseWordDocument(content)
	if err != nil {
		t.Fatalf("Failed to parse document: %v", err)
	}
	
	paragraphs := parser.ExtractParagraphs(doc)
	
	if len(paragraphs) != 2 {
		t.Errorf("Expected 2 paragraphs, got %d", len(paragraphs))
	}
	
	// 验证第一段的格式
	firstParagraph := paragraphs[0]
	if firstParagraph.Text != "这是一个测试文档" {
		t.Errorf("Expected first paragraph text '这是一个测试文档', got '%s'", firstParagraph.Text)
	}
	
	if len(firstParagraph.Runs) != 1 {
		t.Errorf("Expected 1 run in first paragraph, got %d", len(firstParagraph.Runs))
	}
	
	firstRun := firstParagraph.Runs[0]
	if !firstRun.Bold {
		t.Error("Expected first run to be bold")
	}
	
	if firstRun.FontSize != 24 {
		t.Errorf("Expected font size 24, got %d", firstRun.FontSize)
	}
	
	if firstRun.FontName != "Arial" {
		t.Errorf("Expected font name 'Arial', got '%s'", firstRun.FontName)
	}
	
	// 验证第二段的格式
	secondParagraph := paragraphs[1]
	if secondParagraph.Text != "这是第二段，包含斜体文本。" {
		t.Errorf("Expected second paragraph text '这是第二段，包含斜体文本。', got '%s'", secondParagraph.Text)
	}
	
	if len(secondParagraph.Runs) != 1 {
		t.Errorf("Expected 1 run in second paragraph, got %d", len(secondParagraph.Runs))
	}
	
	secondRun := secondParagraph.Runs[0]
	if !secondRun.Italic {
		t.Error("Expected second run to be italic")
	}
	
	if secondRun.FontSize != 12 {
		t.Errorf("Expected font size 12, got %d", secondRun.FontSize)
	}
}

func TestWordMLParserExtractParagraphsWithEmptyDocument(t *testing.T) {
	parser := &WordMLParser{}
	doc := &WordDocument{
		Body: WordBody{
			Paragraphs: []WordParagraph{},
		},
	}
	
	paragraphs := parser.ExtractParagraphs(doc)
	if len(paragraphs) != 0 {
		t.Errorf("Expected 0 paragraphs, got %d", len(paragraphs))
	}
}

func TestWordMLParserExtractTables(t *testing.T) {
	content, err := os.ReadFile("../../tests/test_document.xml")
	if err != nil {
		t.Fatalf("Failed to read test document: %v", err)
	}
	
	parser := &WordMLParser{}
	doc, err := parser.ParseWordDocument(content)
	if err != nil {
		t.Fatalf("Failed to parse document: %v", err)
	}
	
	tables := parser.ExtractTables(doc)
	
	if len(tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(tables))
	}
	
	table := tables[0]
	if len(table.Rows) != 2 {
		t.Errorf("Expected 2 rows in table, got %d", len(table.Rows))
	}
	
	if len(table.Rows[0].Cells) != 2 {
		t.Errorf("Expected 2 cells in first row, got %d", len(table.Rows[0].Cells))
	}
	
	// 验证表格内容
	firstCell := table.Rows[0].Cells[0]
	if firstCell.Text != "表头1" {
		t.Errorf("Expected first cell text '表头1', got '%s'", firstCell.Text)
	}
	
	secondCell := table.Rows[0].Cells[1]
	if secondCell.Text != "表头2" {
		t.Errorf("Expected second cell text '表头2', got '%s'", secondCell.Text)
	}
}

func TestWordMLParserExtractTablesWithEmptyDocument(t *testing.T) {
	parser := &WordMLParser{}
	doc := &WordDocument{
		Body: WordBody{
			Tables: []WordTable{},
		},
	}
	
	tables := parser.ExtractTables(doc)
	if len(tables) != 0 {
		t.Errorf("Expected 0 tables, got %d", len(tables))
	}
}

func TestWordMLParserExtractParagraphText(t *testing.T) {
	parser := &WordMLParser{}
	
	paragraph := WordParagraph{
		Runs: []WordRun{
			{
				Text: &WordText{Content: "Hello"},
			},
			{
				Text: &WordText{Content: " World"},
			},
		},
	}
	
	text := parser.extractParagraphText(paragraph)
	expected := "Hello World"
	if text != expected {
		t.Errorf("Expected text '%s', got '%s'", expected, text)
	}
}

func TestWordMLParserExtractParagraphTextWithNilRuns(t *testing.T) {
	parser := &WordMLParser{}
	
	paragraph := WordParagraph{
		Runs: nil,
	}
	
	text := parser.extractParagraphText(paragraph)
	if text != "" {
		t.Errorf("Expected empty text, got '%s'", text)
	}
}

func TestWordMLParserExtractParagraphTextWithEmptyRuns(t *testing.T) {
	parser := &WordMLParser{}
	
	paragraph := WordParagraph{
		Runs: []WordRun{},
	}
	
	text := parser.extractParagraphText(paragraph)
	if text != "" {
		t.Errorf("Expected empty text, got '%s'", text)
	}
}

func TestWordMLParserConvertRun(t *testing.T) {
	parser := &WordMLParser{}
	
	wordRun := WordRun{
		Text: &WordText{Content: "Test text"},
		Properties: &RunProps{
			Bold:   &types.Bold{Val: "true"},
			Italic: &types.Italic{Val: "true"},
			Size:   &types.Size{Val: "16"},
			Font:   &types.Font{Ascii: "Arial"},
			Color:  &types.Color{Val: "FF0000"},
		},
	}
	
	run := parser.convertRun(wordRun)
	
	if run.Text != "Test text" {
		t.Errorf("Expected text 'Test text', got '%s'", run.Text)
	}
	
	if !run.Bold {
		t.Error("Expected Bold to be true")
	}
	
	if !run.Italic {
		t.Error("Expected Italic to be true")
	}
	
	if run.FontSize != 16 {
		t.Errorf("Expected font size 16, got %d", run.FontSize)
	}
	
	if run.FontName != "Arial" {
		t.Errorf("Expected font name 'Arial', got '%s'", run.FontName)
	}
	
	if run.Color != "FF0000" {
		t.Errorf("Expected color 'FF0000', got '%s'", run.Color)
	}
}

func TestWordMLParserConvertRunWithNilProperties(t *testing.T) {
	parser := &WordMLParser{}
	
	wordRun := WordRun{
		Text:       &WordText{Content: "Test text"},
		Properties: nil,
	}
	
	run := parser.convertRun(wordRun)
	
	if run.Text != "Test text" {
		t.Errorf("Expected text 'Test text', got '%s'", run.Text)
	}
	
	if run.Bold {
		t.Error("Expected Bold to be false")
	}
	
	if run.Italic {
		t.Error("Expected Italic to be false")
	}
}

func TestWordMLParserConvertRunWithNilText(t *testing.T) {
	parser := &WordMLParser{}
	
	wordRun := WordRun{
		Text:       nil,
		Properties: &RunProps{},
	}
	
	run := parser.convertRun(wordRun)
	
	if run.Text != "" {
		t.Errorf("Expected empty text, got '%s'", run.Text)
	}
}

func TestWordMLParserExtractCellText(t *testing.T) {
	parser := &WordMLParser{}
	
	cell := WordTableCell{
		Paragraphs: []WordParagraph{
			{
				Runs: []WordRun{
					{Text: &WordText{Content: "Cell content"}},
				},
			},
		},
	}
	
	text := parser.extractCellText(cell)
	expected := "Cell content"
	if text != expected {
		t.Errorf("Expected text '%s', got '%s'", expected, text)
	}
}

func TestWordMLParserExtractCellTextWithEmptyParagraphs(t *testing.T) {
	parser := &WordMLParser{}
	
	cell := WordTableCell{
		Paragraphs: []WordParagraph{},
	}
	
	text := parser.extractCellText(cell)
	if text != "" {
		t.Errorf("Expected empty text, got '%s'", text)
	}
}

func TestWordMLParserExtractCellTextWithNilParagraphs(t *testing.T) {
	parser := &WordMLParser{}
	
	cell := WordTableCell{
		Paragraphs: nil,
	}
	
	text := parser.extractCellText(cell)
	if text != "" {
		t.Errorf("Expected empty text, got '%s'", text)
	}
}

func TestWordMLParserExtractCellTextWithEmptyRuns(t *testing.T) {
	parser := &WordMLParser{}
	
	cell := WordTableCell{
		Paragraphs: []WordParagraph{
			{
				Runs: []WordRun{},
			},
		},
	}
	
	text := parser.extractCellText(cell)
	if text != "" {
		t.Errorf("Expected empty text, got '%s'", text)
	}
}

func TestWordMLParserExtractCellTextWithNilRuns(t *testing.T) {
	parser := &WordMLParser{}
	
	cell := WordTableCell{
		Paragraphs: []WordParagraph{
			{
				Runs: nil,
			},
		},
	}
	
	text := parser.extractCellText(cell)
	if text != "" {
		t.Errorf("Expected empty text, got '%s'", text)
	}
} 