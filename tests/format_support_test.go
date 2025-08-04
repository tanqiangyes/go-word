package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewFormatSupport(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	if support == nil {
		t.Fatal("Expected FormatSupport to be created")
	}
}

func TestDetectFormat(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 测试不同格式的检测
	testCases := []struct {
		filename string
		expected string
	}{
		{"document.docx", "docx"},
		{"document.doc", "doc"},
		{"document.docm", "docm"},
		{"document.rtf", "rtf"},
		{"document.txt", "unknown"},
	}
	
	for _, tc := range testCases {
		format := support.DetectFormat(tc.filename)
		if format != tc.expected {
			t.Errorf("For filename '%s', expected format '%s', got '%s'", 
				tc.filename, tc.expected, format)
		}
	}
}

func TestConvertFormat(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 测试格式转换
	testCases := []struct {
		sourceFormat string
		targetFormat string
		shouldSucceed bool
	}{
		{"docx", "doc", true},
		{"doc", "docx", true},
		{"docx", "docm", true},
		{"docm", "docx", true},
		{"docx", "rtf", true},
		{"rtf", "docx", true},
		{"txt", "docx", false},
	}
	
	for _, tc := range testCases {
		err := support.ConvertFormat("test."+tc.sourceFormat, "output."+tc.targetFormat)
		if tc.shouldSucceed && err != nil {
			t.Errorf("Expected successful conversion from %s to %s, got error: %v", 
				tc.sourceFormat, tc.targetFormat, err)
		}
		if !tc.shouldSucceed && err == nil {
			t.Errorf("Expected conversion from %s to %s to fail", 
				tc.sourceFormat, tc.targetFormat)
		}
	}
}

func TestCreateRichTextContent(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建富文本内容
	content, err := support.CreateRichTextContent("Test rich text content")
	if err != nil {
		t.Fatalf("Failed to create rich text content: %v", err)
	}
	
	if content == nil {
		t.Fatal("Expected rich text content to be created")
	}
	
	if content.Text != "Test rich text content" {
		t.Errorf("Expected text 'Test rich text content', got '%s'", content.Text)
	}
	
	if content.Formatting == nil {
		t.Error("Expected formatting to be initialized")
	}
}

func TestAddRichTextFormatting(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建富文本内容
	content, err := support.CreateRichTextContent("Test content")
	if err != nil {
		t.Fatalf("Failed to create rich text content: %v", err)
	}
	
	// 添加格式化
	formatting := wordprocessingml.RichTextFormatting{
		FontSize:  14,
		FontBold:  true,
		FontColor: "FF0000",
		Alignment: "center",
	}
	
	err = support.AddRichTextFormatting(content, formatting)
	if err != nil {
		t.Fatalf("Failed to add rich text formatting: %v", err)
	}
	
	// 验证格式化
	if content.Formatting.FontSize != 14 {
		t.Errorf("Expected font size 14, got %d", content.Formatting.FontSize)
	}
	
	if !content.Formatting.FontBold {
		t.Error("Expected font to be bold")
	}
	
	if content.Formatting.FontColor != "FF0000" {
		t.Errorf("Expected font color 'FF0000', got '%s'", content.Formatting.FontColor)
	}
	
	if content.Formatting.Alignment != "center" {
		t.Errorf("Expected alignment 'center', got '%s'", content.Formatting.Alignment)
	}
}

func TestAddHyperlink(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建富文本内容
	content, err := support.CreateRichTextContent("Test content")
	if err != nil {
		t.Fatalf("Failed to create rich text content: %v", err)
	}
	
	// 添加超链接
	hyperlink := wordprocessingml.Hyperlink{
		Text: "Click here",
		URL:  "https://example.com",
		Title: "Example Link",
	}
	
	err = support.AddHyperlink(content, hyperlink)
	if err != nil {
		t.Fatalf("Failed to add hyperlink: %v", err)
	}
	
	// 验证超链接
	if len(content.Hyperlinks) == 0 {
		t.Error("Expected hyperlink to be added")
	}
	
	addedLink := content.Hyperlinks[0]
	if addedLink.Text != "Click here" {
		t.Errorf("Expected link text 'Click here', got '%s'", addedLink.Text)
	}
	
	if addedLink.URL != "https://example.com" {
		t.Errorf("Expected link URL 'https://example.com', got '%s'", addedLink.URL)
	}
}

func TestAddImage(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建富文本内容
	content, err := support.CreateRichTextContent("Test content")
	if err != nil {
		t.Fatalf("Failed to create rich text content: %v", err)
	}
	
	// 添加图片
	image := wordprocessingml.Image{
		Path:      "/path/to/image.jpg",
		Width:     300,
		Height:    200,
		AltText:   "Test Image",
		Alignment: "center",
	}
	
	err = support.AddImage(content, image)
	if err != nil {
		t.Fatalf("Failed to add image: %v", err)
	}
	
	// 验证图片
	if len(content.Images) == 0 {
		t.Error("Expected image to be added")
	}
	
	addedImage := content.Images[0]
	if addedImage.Path != "/path/to/image.jpg" {
		t.Errorf("Expected image path '/path/to/image.jpg', got '%s'", addedImage.Path)
	}
	
	if addedImage.Width != 300 {
		t.Errorf("Expected image width 300, got %d", addedImage.Width)
	}
	
	if addedImage.Height != 200 {
		t.Errorf("Expected image height 200, got %d", addedImage.Height)
	}
}

func TestCreateRichTextTable(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建富文本表格
	table, err := support.CreateRichTextTable(3, 4)
	if err != nil {
		t.Fatalf("Failed to create rich text table: %v", err)
	}
	
	if table == nil {
		t.Fatal("Expected rich text table to be created")
	}
	
	if len(table.Rows) != 3 {
		t.Errorf("Expected 3 rows, got %d", len(table.Rows))
	}
	
	if len(table.Rows[0].Columns) != 4 {
		t.Errorf("Expected 4 columns per row, got %d", len(table.Rows[0].Columns))
	}
}

func TestCreateRichTextList(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建富文本列表
	list, err := support.CreateRichTextList("numbered")
	if err != nil {
		t.Fatalf("Failed to create rich text list: %v", err)
	}
	
	if list == nil {
		t.Fatal("Expected rich text list to be created")
	}
	
	if list.Type != "numbered" {
		t.Errorf("Expected list type 'numbered', got '%s'", list.Type)
	}
}

func TestAddListItem(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建富文本列表
	list, err := support.CreateRichTextList("bulleted")
	if err != nil {
		t.Fatalf("Failed to create rich text list: %v", err)
	}
	
	// 添加列表项
	item := wordprocessingml.RichTextListItem{
		Text:     "First item",
		Level:    0,
		Indent:   0,
		Bullet:   "•",
	}
	
	err = support.AddListItem(list, item)
	if err != nil {
		t.Fatalf("Failed to add list item: %v", err)
	}
	
	// 验证列表项
	if len(list.Items) == 0 {
		t.Error("Expected list item to be added")
	}
	
	addedItem := list.Items[0]
	if addedItem.Text != "First item" {
		t.Errorf("Expected item text 'First item', got '%s'", addedItem.Text)
	}
	
	if addedItem.Level != 0 {
		t.Errorf("Expected item level 0, got %d", addedItem.Level)
	}
}

func TestApplyRichTextFormatting(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建富文本内容
	content, err := support.CreateRichTextContent("Test content")
	if err != nil {
		t.Fatalf("Failed to create rich text content: %v", err)
	}
	
	// 应用格式化
	formatting := wordprocessingml.RichTextFormatting{
		FontSize:    16,
		FontFamily:  "Arial",
		FontColor:   "0000FF",
		Alignment:   "left",
		LineSpacing: 1.5,
		Indentation: 10,
	}
	
	err = support.ApplyRichTextFormatting(content, formatting)
	if err != nil {
		t.Fatalf("Failed to apply rich text formatting: %v", err)
	}
	
	// 验证格式化
	if content.Formatting.FontSize != 16 {
		t.Errorf("Expected font size 16, got %d", content.Formatting.FontSize)
	}
	
	if content.Formatting.FontFamily != "Arial" {
		t.Errorf("Expected font family 'Arial', got '%s'", content.Formatting.FontFamily)
	}
	
	if content.Formatting.FontColor != "0000FF" {
		t.Errorf("Expected font color '0000FF', got '%s'", content.Formatting.FontColor)
	}
	
	if content.Formatting.Alignment != "left" {
		t.Errorf("Expected alignment 'left', got '%s'", content.Formatting.Alignment)
	}
	
	if content.Formatting.LineSpacing != 1.5 {
		t.Errorf("Expected line spacing 1.5, got %f", content.Formatting.LineSpacing)
	}
	
	if content.Formatting.Indentation != 10 {
		t.Errorf("Expected indentation 10, got %d", content.Formatting.Indentation)
	}
}

func TestFontProperties(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建字体
	font := wordprocessingml.Font{
		Name:     "Arial",
		Size:     12,
		Bold:     false,
		Italic:   false,
		Underline: false,
		Color:    "000000",
		Family:   "sans-serif",
	}
	
	// 验证字体属性
	if font.Name != "Arial" {
		t.Errorf("Expected font name 'Arial', got '%s'", font.Name)
	}
	
	if font.Size != 12 {
		t.Errorf("Expected font size 12, got %d", font.Size)
	}
	
	if font.Bold {
		t.Error("Expected font to not be bold")
	}
	
	if font.Italic {
		t.Error("Expected font to not be italic")
	}
	
	if font.Underline {
		t.Error("Expected font to not be underlined")
	}
	
	if font.Color != "000000" {
		t.Errorf("Expected font color '000000', got '%s'", font.Color)
	}
	
	if font.Family != "sans-serif" {
		t.Errorf("Expected font family 'sans-serif', got '%s'", font.Family)
	}
}

func TestParagraphFormat(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建段落格式
	format := wordprocessingml.ParagraphFormat{
		Alignment:    "justify",
		LineSpacing:  1.5,
		Indentation:  20,
		Spacing:      wordprocessingml.SpacingFormat{Before: 6, After: 6},
		Borders:      wordprocessingml.BorderFormat{Style: "solid", Width: 1, Color: "000000"},
		Shading:      wordprocessingml.ShadingFormat{Fill: "F0F0F0", Pattern: "solid"},
	}
	
	// 验证段落格式
	if format.Alignment != "justify" {
		t.Errorf("Expected alignment 'justify', got '%s'", format.Alignment)
	}
	
	if format.LineSpacing != 1.5 {
		t.Errorf("Expected line spacing 1.5, got %f", format.LineSpacing)
	}
	
	if format.Indentation != 20 {
		t.Errorf("Expected indentation 20, got %d", format.Indentation)
	}
	
	if format.Spacing.Before != 6 {
		t.Errorf("Expected spacing before 6, got %d", format.Spacing.Before)
	}
	
	if format.Spacing.After != 6 {
		t.Errorf("Expected spacing after 6, got %d", format.Spacing.After)
	}
	
	if format.Borders.Style != "solid" {
		t.Errorf("Expected border style 'solid', got '%s'", format.Borders.Style)
	}
	
	if format.Borders.Width != 1 {
		t.Errorf("Expected border width 1, got %d", format.Borders.Width)
	}
	
	if format.Shading.Fill != "F0F0F0" {
		t.Errorf("Expected shading fill 'F0F0F0', got '%s'", format.Shading.Fill)
	}
	
	if format.Shading.Pattern != "solid" {
		t.Errorf("Expected shading pattern 'solid', got '%s'", format.Shading.Pattern)
	}
}

func TestListProperties(t *testing.T) {
	support := wordprocessingml.NewFormatSupport()
	
	// 创建列表属性
	properties := wordprocessingml.ListProperties{
		Type:        "bulleted",
		StartNumber: 1,
		IndentLevel: 0,
		BulletChar:  "•",
		NumberFormat: "decimal",
		Restart:     "continuous",
	}
	
	// 验证列表属性
	if properties.Type != "bulleted" {
		t.Errorf("Expected list type 'bulleted', got '%s'", properties.Type)
	}
	
	if properties.StartNumber != 1 {
		t.Errorf("Expected start number 1, got %d", properties.StartNumber)
	}
	
	if properties.IndentLevel != 0 {
		t.Errorf("Expected indent level 0, got %d", properties.IndentLevel)
	}
	
	if properties.BulletChar != "•" {
		t.Errorf("Expected bullet character '•', got '%s'", properties.BulletChar)
	}
	
	if properties.NumberFormat != "decimal" {
		t.Errorf("Expected number format 'decimal', got '%s'", properties.NumberFormat)
	}
	
	if properties.Restart != "continuous" {
		t.Errorf("Expected restart 'continuous', got '%s'", properties.Restart)
	}
}

// 辅助函数
