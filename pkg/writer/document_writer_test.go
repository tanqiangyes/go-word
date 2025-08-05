package writer

import (
	"testing"

	"github.com/tanqiangyes/go-word/pkg/types"
)

func TestDocumentWriterCreateNewDocument(t *testing.T) {
	writer := NewDocumentWriter()
	
	if writer == nil {
		t.Fatal("Expected document writer to be created")
	}
	
	// 创建新文档
	err := writer.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}
}

func TestDocumentWriterAddParagraph(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加段落
	err := writer.AddParagraph("Test paragraph", "Normal")
	if err != nil {
		t.Fatalf("Failed to add paragraph: %v", err)
	}

	// 验证段落已添加
	mainPart := writer.document.GetMainPart()
	if len(mainPart.Content.Paragraphs) != 1 {
		t.Errorf("Expected 1 paragraph, got %d", len(mainPart.Content.Paragraphs))
	}

	if mainPart.Content.Paragraphs[0].Text != "Test paragraph" {
		t.Errorf("Expected paragraph text 'Test paragraph', got '%s'", mainPart.Content.Paragraphs[0].Text)
	}
}

func TestDocumentWriterAddParagraphWithoutInitialization(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试未初始化文档的情况
	err := writer.AddParagraph("Test paragraph", "Normal")
	if err == nil {
		t.Error("Expected error when document not initialized")
	}
}

func TestDocumentWriterAddTable(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加表格
	rows := [][]string{
		{"Header 1", "Header 2"},
		{"Data 1", "Data 2"},
	}

	err := writer.AddTable(rows)
	if err != nil {
		t.Fatalf("Failed to add table: %v", err)
	}

	// 验证表格已添加
	mainPart := writer.document.GetMainPart()
	if len(mainPart.Content.Tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(mainPart.Content.Tables))
	}

	if len(mainPart.Content.Tables[0].Rows) != 2 {
		t.Errorf("Expected 2 rows, got %d", len(mainPart.Content.Tables[0].Rows))
	}
}

func TestDocumentWriterAddTableWithoutInitialization(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试未初始化文档的情况
	rows := [][]string{{"Header", "Data"}}
	err := writer.AddTable(rows)
	if err == nil {
		t.Error("Expected error when document not initialized")
	}
}

func TestDocumentWriterAddFormattedParagraph(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加格式化段落
	runs := []types.Run{
		{
			Text:      "Bold text",
			Bold:      true,
			FontSize:  16,
			FontName:  "Arial",
		},
	}

	err := writer.AddFormattedParagraph("Formatted paragraph", "Normal", runs)
	if err != nil {
		t.Fatalf("Failed to add formatted paragraph: %v", err)
	}

	// 验证格式化段落已添加
	mainPart := writer.document.GetMainPart()
	if len(mainPart.Content.Paragraphs) != 1 {
		t.Errorf("Expected 1 paragraph, got %d", len(mainPart.Content.Paragraphs))
	}

	if len(mainPart.Content.Paragraphs[0].Runs) != 1 {
		t.Errorf("Expected 1 run, got %d", len(mainPart.Content.Paragraphs[0].Runs))
	}

	run := mainPart.Content.Paragraphs[0].Runs[0]
	if !run.Bold {
		t.Error("Expected run to be bold")
	}

	if run.FontSize != 16 {
		t.Errorf("Expected font size 16, got %d", run.FontSize)
	}
}

func TestDocumentWriterAddFormattedParagraphWithoutInitialization(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试未初始化文档的情况
	runs := []types.Run{{Text: "Test"}}
	err := writer.AddFormattedParagraph("Test", "Normal", runs)
	if err == nil {
		t.Error("Expected error when document not initialized")
	}
}

func TestDocumentWriterReplaceText(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加一些内容
	writer.AddParagraph("Original text", "Normal")
	writer.AddParagraph("Another paragraph", "Normal")

	// 替换文本
	err := writer.ReplaceText("Original", "Modified")
	if err != nil {
		t.Fatalf("Failed to replace text: %v", err)
	}

	// 验证文本已替换
	mainPart := writer.document.GetMainPart()
	if mainPart.Content.Paragraphs[0].Text != "Modified text" {
		t.Errorf("Expected 'Modified text', got '%s'", mainPart.Content.Paragraphs[0].Text)
	}
}

func TestDocumentWriterReplaceTextWithoutInitialization(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试未初始化文档的情况
	err := writer.ReplaceText("Original", "Modified")
	if err == nil {
		t.Error("Expected error when document not initialized")
	}
}

func TestDocumentWriterSave(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加一些内容
	writer.AddParagraph("Test paragraph", "Normal")

	// 保存文档
	filename := "test_save.docx"
	err := writer.Save(filename)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}
}

func TestDocumentWriterSaveWithoutInitialization(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试未初始化文档的情况
	err := writer.Save("test.docx")
	if err == nil {
		t.Error("Expected error when document not initialized")
	}
}

func TestDocumentWriterSetParagraphStyle(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加段落
	writer.AddParagraph("Test paragraph", "Normal")

	// 设置段落样式
	err := writer.SetParagraphStyle(0, "Heading1")
	if err != nil {
		t.Fatalf("Failed to set paragraph style: %v", err)
	}

	// 验证样式已设置
	mainPart := writer.document.GetMainPart()
	if mainPart.Content.Paragraphs[0].Style != "Heading1" {
		t.Errorf("Expected style 'Heading1', got '%s'", mainPart.Content.Paragraphs[0].Style)
	}
}

func TestDocumentWriterSetParagraphStyleWithoutInitialization(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试未初始化文档的情况
	err := writer.SetParagraphStyle(0, "Heading1")
	if err == nil {
		t.Error("Expected error when document not initialized")
	}
}

func TestDocumentWriterSetParagraphStyleWithInvalidIndex(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 测试无效的段落索引
	err := writer.SetParagraphStyle(999, "Invalid")
	if err == nil {
		t.Error("Expected error for invalid paragraph index")
	}
}

func TestDocumentWriterSetRunFormatting(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加段落
	writer.AddParagraph("Test paragraph", "Normal")

	// 设置运行格式
	formatting := types.Run{
		Bold:      true,
		Italic:    true,
		FontSize:  18,
		FontName:  "Times New Roman",
		Color:     "0000FF",
	}

	err := writer.SetRunFormatting(0, 0, formatting)
	if err != nil {
		t.Fatalf("Failed to set run formatting: %v", err)
	}

	// 验证格式已设置
	mainPart := writer.document.GetMainPart()
	run := mainPart.Content.Paragraphs[0].Runs[0]
	if !run.Bold {
		t.Error("Expected Bold to be true")
	}

	if !run.Italic {
		t.Error("Expected Italic to be true")
	}

	if run.FontSize != 18 {
		t.Errorf("Expected font size 18, got %d", run.FontSize)
	}

	if run.FontName != "Times New Roman" {
		t.Errorf("Expected font name 'Times New Roman', got '%s'", run.FontName)
	}

	if run.Color != "0000FF" {
		t.Errorf("Expected color '0000FF', got '%s'", run.Color)
	}
}

func TestDocumentWriterSetRunFormattingWithoutInitialization(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试未初始化文档的情况
	formatting := types.Run{Bold: true}
	err := writer.SetRunFormatting(0, 0, formatting)
	if err == nil {
		t.Error("Expected error when document not initialized")
	}
}

func TestDocumentWriterSetRunFormattingWithInvalidIndex(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 测试无效的运行索引
	err := writer.SetRunFormatting(999, 999, types.Run{})
	if err == nil {
		t.Error("Expected error for invalid run index")
	}
}

func TestDocumentWriterErrorHandling(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试无效的段落索引
	err := writer.SetParagraphStyle(999, "Invalid")
	if err == nil {
		t.Error("Expected error for invalid paragraph index")
	}

	// 测试无效的运行索引
	err = writer.SetRunFormatting(999, 999, types.Run{})
	if err == nil {
		t.Error("Expected error for invalid run index")
	}
}

func TestDocumentWriterMultipleParagraphs(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加多个段落
	writer.AddParagraph("First paragraph", "Normal")
	writer.AddParagraph("Second paragraph", "Normal")
	writer.AddParagraph("Third paragraph", "Normal")

	// 验证所有段落都已添加
	mainPart := writer.document.GetMainPart()
	if len(mainPart.Content.Paragraphs) != 3 {
		t.Errorf("Expected 3 paragraphs, got %d", len(mainPart.Content.Paragraphs))
	}

	// 验证段落内容
	expectedTexts := []string{"First paragraph", "Second paragraph", "Third paragraph"}
	for i, expected := range expectedTexts {
		if mainPart.Content.Paragraphs[i].Text != expected {
			t.Errorf("Expected paragraph %d text '%s', got '%s'", i, expected, mainPart.Content.Paragraphs[i].Text)
		}
	}
}

func TestDocumentWriterMultipleTables(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加多个表格
	table1 := [][]string{{"Table 1"}}
	table2 := [][]string{{"Table 2"}}

	writer.AddTable(table1)
	writer.AddTable(table2)

	// 验证所有表格都已添加
	mainPart := writer.document.GetMainPart()
	if len(mainPart.Content.Tables) != 2 {
		t.Errorf("Expected 2 tables, got %d", len(mainPart.Content.Tables))
	}
}

func TestDocumentWriterOpenForModification(t *testing.T) {
	writer := NewDocumentWriter()
	
	// 测试打开不存在的文档
	err := writer.OpenForModification("nonexistent.docx")
	if err == nil {
		t.Error("Expected error when opening nonexistent document")
	}
}

func TestDocumentWriterGenerateXML(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 添加一些内容
	writer.AddParagraph("Test paragraph", "Normal")

	// 测试XML生成
	xmlData, err := writer.generateDocumentXML()
	if err != nil {
		t.Fatalf("Failed to generate document XML: %v", err)
	}

	if len(xmlData) == 0 {
		t.Error("Expected XML data to not be empty")
	}
}

func TestDocumentWriterGenerateContentTypes(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 测试ContentTypes XML生成
	contentTypesXML := writer.generateContentTypesXML()
	if len(contentTypesXML) == 0 {
		t.Error("Expected ContentTypes XML to not be empty")
	}
}

func TestDocumentWriterGenerateRootRels(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 测试Root Rels XML生成
	rootRelsXML := writer.generateRootRelsXML()
	if len(rootRelsXML) == 0 {
		t.Error("Expected Root Rels XML to not be empty")
	}
}

func TestDocumentWriterGenerateDocumentRels(t *testing.T) {
	writer := NewDocumentWriter()
	writer.CreateNewDocument()

	// 测试Document Rels XML生成
	documentRelsXML := writer.generateDocumentRelsXML()
	if len(documentRelsXML) == 0 {
		t.Error("Expected Document Rels XML to not be empty")
	}
} 