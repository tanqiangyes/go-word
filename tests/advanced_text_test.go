package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewAdvancedTextSystem(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	if system == nil {
		t.Fatal("Expected AdvancedTextSystem to be created")
	}
	
	if system.TextManager == nil {
		t.Error("Expected TextManager to be initialized")
	}
	
	if system.TextStyles == nil {
		t.Error("Expected TextStyles to be initialized")
	}
	
	if system.TextEffects == nil {
		t.Error("Expected TextEffects to be initialized")
	}
	
	if system.TextValidator == nil {
		t.Error("Expected TextValidator to be initialized")
	}
}

func TestCreateAdvancedText(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建文本属性
	properties := wordprocessingml.AdvancedTextProperties{
		FontSize:    12,
		FontFamily:  "Arial",
		FontColor:   "000000",
		Alignment:   "left",
		LineSpacing: 1.5,
		Indentation: 0,
		Bold:        false,
		Italic:      false,
		Underline:   false,
	}
	
	// 创建高级文本
	text, err := system.CreateAdvancedText("Test content", properties)
	if err != nil {
		t.Fatalf("Failed to create advanced text: %v", err)
	}
	
	if text == nil {
		t.Fatal("Expected text to be created")
	}
	
	if text.Content != "Test content" {
		t.Errorf("Expected content 'Test content', got '%s'", text.Content)
	}
	
	if text.Properties.FontSize != 12 {
		t.Errorf("Expected font size 12, got %d", text.Properties.FontSize)
	}
	
	if text.Properties.FontFamily != "Arial" {
		t.Errorf("Expected font family 'Arial', got '%s'", text.Properties.FontFamily)
	}
}

func TestApplyTextEffect(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建文本
	properties := wordprocessingml.AdvancedTextProperties{
		FontSize:  12,
		FontColor: "000000",
	}
	
	text, err := system.CreateAdvancedText("Test text", properties)
	if err != nil {
		t.Fatalf("Failed to create text: %v", err)
	}
	
	// 应用文本效果
	effect := wordprocessingml.TextEffect{
		Type:        "shadow",
		Color:       "666666",
		BlurRadius:  2,
		OffsetX:     1,
		OffsetY:     1,
		Opacity:     0.5,
	}
	
	err = system.ApplyTextEffect(text, effect)
	if err != nil {
		t.Fatalf("Failed to apply text effect: %v", err)
	}
	
	// 验证效果是否被应用
	if len(text.Effects) == 0 {
		t.Error("Expected text effects to be applied")
	}
	
	appliedEffect := text.Effects[0]
	if appliedEffect.Type != "shadow" {
		t.Errorf("Expected effect type 'shadow', got '%s'", appliedEffect.Type)
	}
	
	if appliedEffect.Color != "666666" {
		t.Errorf("Expected effect color '666666', got '%s'", appliedEffect.Color)
	}
}

func TestGetTextSummary(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建文本
	properties := wordprocessingml.AdvancedTextProperties{
		FontSize:    14,
		FontFamily:  "Times New Roman",
		FontColor:   "000000",
		Alignment:   "center",
		LineSpacing: 1.2,
		Bold:        true,
		Italic:      false,
	}
	
	text, err := system.CreateAdvancedText("This is a test text with multiple words", properties)
	if err != nil {
		t.Fatalf("Failed to create text: %v", err)
	}
	
	// 应用效果
	effect := wordprocessingml.TextEffect{
		Type:  "highlight",
		Color: "FFFF00",
	}
	system.ApplyTextEffect(text, effect)
	
	summary := system.GetTextSummary(text)
	
	if summary == "" {
		t.Error("Expected non-empty text summary")
	}
	
	// 检查摘要是否包含预期的文本信息
	expectedInfo := []string{"14", "Times New Roman", "center", "bold", "highlight", "8 words"}
	for _, expected := range expectedInfo {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

func TestTextValidation(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建文本
	properties := wordprocessingml.AdvancedTextProperties{
		FontSize:  12,
		FontColor: "000000",
	}
	
	text, err := system.CreateAdvancedText("This is a test text with some spelling errors like 'recieve' and 'seperate'.", properties)
	if err != nil {
		t.Fatalf("Failed to create text: %v", err)
	}
	
	// 验证文本
	results := system.TextValidator.ValidateText(text)
	
	if len(results) == 0 {
		t.Error("Expected validation results")
	}
	
	// 检查是否有验证错误
	hasErrors := false
	for _, result := range results {
		if result.Severity == "error" {
			hasErrors = true
			break
		}
	}
	
	// 对于包含拼写错误的文本，应该检测到错误
	if !hasErrors {
		t.Error("Expected text validation to detect spelling errors")
	}
}

func TestFormattedTextContent(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建格式化文本内容
	formattedText := wordprocessingml.FormattedTextContent{
		Text: "This is formatted text",
		Runs: []wordprocessingml.AdvancedTextRun{
			{
				Text: "This is ",
				Properties: wordprocessingml.RunProperties{
					FontSize:  12,
					FontBold:  false,
					FontColor: "000000",
				},
			},
			{
				Text: "formatted",
				Properties: wordprocessingml.RunProperties{
					FontSize:  14,
					FontBold:  true,
					FontColor: "FF0000",
				},
			},
			{
				Text: " text",
				Properties: wordprocessingml.RunProperties{
					FontSize:  12,
					FontBold:  false,
					FontColor: "000000",
				},
			},
		},
	}
	
	// 验证格式化文本
	if formattedText.Text != "This is formatted text" {
		t.Errorf("Expected text 'This is formatted text', got '%s'", formattedText.Text)
	}
	
	if len(formattedText.Runs) != 3 {
		t.Errorf("Expected 3 runs, got %d", len(formattedText.Runs))
	}
	
	// 验证第二个run的格式
	boldRun := formattedText.Runs[1]
	if !boldRun.Properties.FontBold {
		t.Error("Expected second run to be bold")
	}
	
	if boldRun.Properties.FontSize != 14 {
		t.Errorf("Expected second run font size 14, got %d", boldRun.Properties.FontSize)
	}
	
	if boldRun.Properties.FontColor != "FF0000" {
		t.Errorf("Expected second run color 'FF0000', got '%s'", boldRun.Properties.FontColor)
	}
}

func TestAdvancedParagraph(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建高级段落
	paragraph := wordprocessingml.AdvancedParagraph{
		Content: "This is a test paragraph with multiple sentences. It contains various formatting options.",
		Properties: wordprocessingml.ParagraphProperties{
			Alignment:    "justify",
			LineSpacing:  1.5,
			Indentation:  10,
			Spacing:      wordprocessingml.ParagraphSpacing{Before: 6, After: 6},
			Style:        "Normal",
		},
		Runs: []wordprocessingml.AdvancedTextRun{
			{
				Text: "This is a test paragraph ",
				Properties: wordprocessingml.RunProperties{
					FontSize:  12,
					FontBold:  false,
				},
			},
			{
				Text: "with multiple sentences",
				Properties: wordprocessingml.RunProperties{
					FontSize:  12,
					FontBold:  true,
					FontColor: "0000FF",
				},
			},
		},
	}
	
	// 验证段落属性
	if paragraph.Properties.Alignment != "justify" {
		t.Errorf("Expected alignment 'justify', got '%s'", paragraph.Properties.Alignment)
	}
	
	if paragraph.Properties.LineSpacing != 1.5 {
		t.Errorf("Expected line spacing 1.5, got %f", paragraph.Properties.LineSpacing)
	}
	
	if paragraph.Properties.Indentation != 10 {
		t.Errorf("Expected indentation 10, got %d", paragraph.Properties.Indentation)
	}
	
	// 验证runs
	if len(paragraph.Runs) != 2 {
		t.Errorf("Expected 2 runs, got %d", len(paragraph.Runs))
	}
	
	boldRun := paragraph.Runs[1]
	if !boldRun.Properties.FontBold {
		t.Error("Expected second run to be bold")
	}
	
	if boldRun.Properties.FontColor != "0000FF" {
		t.Errorf("Expected second run color '0000FF', got '%s'", boldRun.Properties.FontColor)
	}
}

func TestTextEffects(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建文本效果
	effects := []wordprocessingml.TextEffect{
		{
			Type:    "shadow",
			Color:   "666666",
			OffsetX: 2,
			OffsetY: 2,
			BlurRadius: 3,
		},
		{
			Type:    "glow",
			Color:   "FFFF00",
			Radius:  5,
			Opacity: 0.7,
		},
		{
			Type:    "outline",
			Color:   "000000",
			Width:   1,
		},
	}
	
	// 创建文本并应用效果
	properties := wordprocessingml.AdvancedTextProperties{
		FontSize:  16,
		FontColor: "FF0000",
	}
	
	text, err := system.CreateAdvancedText("Text with effects", properties)
	if err != nil {
		t.Fatalf("Failed to create text: %v", err)
	}
	
	// 应用多个效果
	for _, effect := range effects {
		err = system.ApplyTextEffect(text, effect)
		if err != nil {
			t.Fatalf("Failed to apply effect: %v", err)
		}
	}
	
	// 验证效果数量
	if len(text.Effects) != 3 {
		t.Errorf("Expected 3 effects, got %d", len(text.Effects))
	}
	
	// 验证第一个效果
	shadowEffect := text.Effects[0]
	if shadowEffect.Type != "shadow" {
		t.Errorf("Expected first effect type 'shadow', got '%s'", shadowEffect.Type)
	}
	
	if shadowEffect.OffsetX != 2 {
		t.Errorf("Expected shadow offset X 2, got %d", shadowEffect.OffsetX)
	}
}

func TestTextStatistics(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建文本
	properties := wordprocessingml.AdvancedTextProperties{
		FontSize:  12,
		FontColor: "000000",
	}
	
	text, err := system.CreateAdvancedText("This is a test text with multiple words and sentences. It contains various characters and formatting.", properties)
	if err != nil {
		t.Fatalf("Failed to create text: %v", err)
	}
	
	// 获取文本统计信息
	stats := system.TextManager.GetTextStatistics(text)
	
	if stats.WordCount != 15 {
		t.Errorf("Expected word count 15, got %d", stats.WordCount)
	}
	
	if stats.CharacterCount != 108 {
		t.Errorf("Expected character count 108, got %d", stats.CharacterCount)
	}
	
	if stats.SentenceCount != 2 {
		t.Errorf("Expected sentence count 2, got %d", stats.SentenceCount)
	}
	
	if stats.ParagraphCount != 1 {
		t.Errorf("Expected paragraph count 1, got %d", stats.ParagraphCount)
	}
}

func TestTextStyles(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建文本样式
	style := wordprocessingml.AdvancedTextStyle{
		ID:          "TestStyle",
		Name:        "Test Text Style",
		FontSize:    14,
		FontFamily:  "Arial",
		FontColor:   "000000",
		Alignment:   "left",
		LineSpacing: 1.2,
		Bold:        false,
		Italic:      false,
		Underline:   false,
	}
	
	// 添加样式
	system.TextStyles["TestStyle"] = &style
	
	// 验证样式是否被添加
	found := system.TextStyles["TestStyle"]
	if found == nil {
		t.Error("Expected to find added text style")
	}
	
	if found.ID != "TestStyle" {
		t.Errorf("Expected style ID 'TestStyle', got '%s'", found.ID)
	}
	
	if found.FontSize != 14 {
		t.Errorf("Expected font size 14, got %d", found.FontSize)
	}
}

func TestTextValidationRules(t *testing.T) {
	system := wordprocessingml.NewAdvancedTextSystem()
	
	// 创建验证规则
	rule := wordprocessingml.TextValidationRule{
		ID:          "SpellingCheck",
		Name:        "Spelling Check",
		Description: "Check for spelling errors",
		Severity:    "warning",
		Enabled:     true,
		Pattern:     "\\b\\w+\\b",
		Message:     "Possible spelling error: {word}",
	}
	
	// 添加验证规则
	system.TextValidator.AddRule(rule)
	
	// 验证规则是否被添加
	rules := system.TextValidator.GetRules()
	if len(rules) == 0 {
		t.Error("Expected validation rules to be added")
	}
	
	found := false
	for _, r := range rules {
		if r.ID == "SpellingCheck" {
			found = true
			break
		}
	}
	
	if !found {
		t.Error("Expected to find added validation rule")
	}
}

// 辅助函数
