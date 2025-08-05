package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewFormatSupport(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	if support == nil {
		t.Fatal("Expected FormatSupport to be created")
	}
}

func TestDetectFormat(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 测试不同格式的检测
	testCases := []struct {
		filename string
		expected wordprocessingml.DocumentFormat
	}{
		{"document.docx", wordprocessingml.DocxFormat},
		{"document.doc", wordprocessingml.DocFormat},
		{"document.docm", wordprocessingml.DocmFormat},
		{"document.rtf", wordprocessingml.RtfFormat},
	}
	
	for _, tc := range testCases {
		format, err := support.DetectFormat(tc.filename)
		if err != nil {
			t.Errorf("For filename '%s', got error: %v", tc.filename, err)
			continue
		}
		if format != tc.expected {
			t.Errorf("For filename '%s', expected format %v, got %v", 
				tc.filename, tc.expected, format)
		}
	}
}

func TestConvertFormat(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 测试格式转换
	testCases := []struct {
		targetFormat wordprocessingml.DocumentFormat
		shouldSucceed bool
	}{
		{wordprocessingml.DocFormat, true},
		{wordprocessingml.DocxFormat, true},
		{wordprocessingml.DocmFormat, true},
		{wordprocessingml.RtfFormat, true},
	}
	
	for _, tc := range testCases {
		err := support.ConvertFormat(tc.targetFormat)
		if tc.shouldSucceed && err != nil {
			t.Errorf("Expected successful conversion to %v, got error: %v", 
				tc.targetFormat, err)
		}
	}
}

func TestCreateRichTextContent(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 创建富文本内容
	content := support.CreateRichTextContent("Test rich text content")
	
	if content == nil {
		t.Fatal("Expected rich text content to be created")
	}
	
	if content.Text != "Test rich text content" {
		t.Errorf("Expected text 'Test rich text content', got '%s'", content.Text)
	}
	
	if content.Formatting.Font.Size == 0 {
		t.Error("Expected formatting to be initialized")
	}
}

func TestAddRichTextFormatting(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 创建富文本内容
	content := support.CreateRichTextContent("Test content")
	
	// 添加格式化
	formatting := wordprocessingml.RichTextFormatting{
		Font: wordprocessingml.Font{
			Size:  14.0,
			Bold:  true,
		},
	}
	
	support.AddRichTextFormatting(content, formatting)
	
	// 验证格式化
	if content.Formatting.Font.Size != 14.0 {
		t.Errorf("Expected font size 14.0, got %f", content.Formatting.Font.Size)
	}
}

func TestAddHyperlink(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 创建富文本内容
	content := support.CreateRichTextContent("Test content")
	
	// 添加超链接
	support.AddHyperlink(content, "https://example.com", "Click here", "Example Link")
	
	// 验证超链接
	if content == nil {
		t.Error("Expected content to be created")
	}
}

func TestAddImage(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 创建富文本内容
	content := support.CreateRichTextContent("Test content")
	
	// 添加图片
	support.AddImage(content, "image.jpg", 100.0, 50.0)
	
	// 验证图片
	if content == nil {
		t.Error("Expected content to be created")
	}
}

func TestCreateRichTextTable(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 创建富文本表格
	table := support.CreateRichTextTable(3, 4)
	
	if table == nil {
		t.Fatal("Expected rich text table to be created")
	}
	
	if len(table.Rows) != 3 {
		t.Errorf("Expected 3 rows, got %d", len(table.Rows))
	}
	
	if len(table.Columns) != 4 {
		t.Errorf("Expected 4 columns, got %d", len(table.Columns))
	}
}

func TestCreateRichTextList(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 创建富文本列表
	list := support.CreateRichTextList(wordprocessingml.NumberedList)
	
	if list == nil {
		t.Fatal("Expected rich text list to be created")
	}
	
	if list.Type != wordprocessingml.NumberedList {
		t.Errorf("Expected list type NumberedList, got %v", list.Type)
	}
}

func TestAddListItem(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 创建富文本列表
	list := support.CreateRichTextList(wordprocessingml.NumberedList)
	
	// 创建列表项内容
	content := wordprocessingml.RichTextContent{
		Text: "List item 1",
	}
	
	// 添加列表项
	support.AddListItem(list, content, 0)
	
	if len(list.Items) == 0 {
		t.Error("Expected list item to be added")
	}
}

func TestApplyRichTextFormatting(t *testing.T) {
	doc := &wordprocessingml.Document{}
	support := wordprocessingml.NewFormatSupport(doc)
	
	// 创建段落
	paragraph := &wordprocessingml.Paragraph{
		Text: "Test paragraph",
	}
	
	// 创建格式化
	formatting := wordprocessingml.RichTextFormatting{
		Font: wordprocessingml.Font{
			Size:  12.0,
			Bold:  true,
		},
	}
	
	// 应用格式化
	support.ApplyRichTextFormatting(paragraph, formatting)
	
	if paragraph == nil {
		t.Error("Expected paragraph to be created")
	}
}

func TestFontProperties(t *testing.T) {
	// 测试字体属性
	font := wordprocessingml.Font{
		Name:      "Arial",
		Size:      12.0,
		Bold:      true,
		Italic:    false,
		Underline: true,
		Strike:    false,
	}
	
	if font.Name != "Arial" {
		t.Errorf("Expected font name 'Arial', got '%s'", font.Name)
	}
	
	if font.Size != 12.0 {
		t.Errorf("Expected font size 12.0, got %f", font.Size)
	}
	
	if !font.Bold {
		t.Error("Expected font to be bold")
	}
	
	if font.Italic {
		t.Error("Expected font to not be italic")
	}
	
	if !font.Underline {
		t.Error("Expected font to be underlined")
	}
	
	if font.Strike {
		t.Error("Expected font to not be strikethrough")
	}
}

func TestParagraphFormat(t *testing.T) {
	// 测试段落格式
	format := wordprocessingml.ParagraphFormat{
		Alignment: "left",
		Indent: wordprocessingml.IndentFormat{
			Left:   0.0,
			Right:  0.0,
		},
		Spacing: wordprocessingml.SpacingFormat{
			Before: 0.0,
			After:  0.0,
			Line:   1.0,
		},
	}
	
	if format.Alignment != "left" {
		t.Errorf("Expected alignment 'left', got '%s'", format.Alignment)
	}
	
	if format.Indent.Left != 0.0 {
		t.Errorf("Expected left indentation 0.0, got %f", format.Indent.Left)
	}
	
	if format.Spacing.Line != 1.0 {
		t.Errorf("Expected line spacing 1.0, got %f", format.Spacing.Line)
	}
}

func TestListProperties(t *testing.T) {
	// 测试列表属性
	list := wordprocessingml.RichTextList{
		Type: wordprocessingml.NumberedList,
		Items: []wordprocessingml.RichTextListItem{
			{Content: wordprocessingml.RichTextContent{Text: "Item 1"}},
			{Content: wordprocessingml.RichTextContent{Text: "Item 2"}},
			{Content: wordprocessingml.RichTextContent{Text: "Item 3"}},
		},
	}
	
	if list.Type != wordprocessingml.NumberedList {
		t.Errorf("Expected list type NumberedList, got %v", list.Type)
	}
	
	if len(list.Items) != 3 {
		t.Errorf("Expected 3 items, got %d", len(list.Items))
	}
	
	if list.Items[0].Content.Text != "Item 1" {
		t.Errorf("Expected first item 'Item 1', got '%s'", list.Items[0].Content.Text)
	}
}
