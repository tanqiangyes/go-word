package tests

import (
	"os"
	"testing"

	"github.com/tanqiangyes/go-word/pkg/parser"
	"github.com/tanqiangyes/go-word/pkg/types"
)

func TestWordMLParserParseDocument(t *testing.T) {
	// 读取测试文档
	content, err := os.ReadFile("tests/test_document.xml")
	if err != nil {
		t.Fatalf("Failed to read test document: %v", err)
	}
	
	parser := &parser.WordMLParser{}
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

func TestWordMLParserExtractText(t *testing.T) {
	content, err := os.ReadFile("tests/test_document.xml")
	if err != nil {
		t.Fatalf("Failed to read test document: %v", err)
	}
	
	parser := &parser.WordMLParser{}
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

func TestWordMLParserExtractParagraphs(t *testing.T) {
	content, err := os.ReadFile("tests/test_document.xml")
	if err != nil {
		t.Fatalf("Failed to read test document: %v", err)
	}
	
	parser := &parser.WordMLParser{}
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
	
	secondRun := secondParagraph.Runs[0]
	if !secondRun.Italic {
		t.Error("Expected second run to be italic")
	}
	
	if secondRun.FontSize != 12 {
		t.Errorf("Expected font size 12, got %d", secondRun.FontSize)
	}
}

func TestWordMLParserExtractTables(t *testing.T) {
	content, err := os.ReadFile("tests/test_document.xml")
	if err != nil {
		t.Fatalf("Failed to read test document: %v", err)
	}
	
	parser := &parser.WordMLParser{}
	doc, err := parser.ParseWordDocument(content)
	if err != nil {
		t.Fatalf("Failed to parse document: %v", err)
	}
	
	tables := parser.ExtractTables(doc)
	
	if len(tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(tables))
	}
	
	table := tables[0]
	if len(table.Rows) != 1 {
		t.Errorf("Expected 1 row in table, got %d", len(table.Rows))
	}
	
	row := table.Rows[0]
	if len(row.Cells) != 2 {
		t.Errorf("Expected 2 cells in row, got %d", len(row.Cells))
	}
	
	if table.Columns != 2 {
		t.Errorf("Expected 2 columns in table, got %d", table.Columns)
	}
} 