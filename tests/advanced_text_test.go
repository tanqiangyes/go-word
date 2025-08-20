package tests

import (
	"testing"

	"github.com/tanqiangyes/go-word/pkg/word"
)

func TestNewAdvancedTextSystem(t *testing.T) {
	system := word.NewAdvancedTextSystem()

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
	system := word.NewAdvancedTextSystem()

	// 创建高级文本
	text := system.CreateAdvancedText("Test Text", "This is a test text content")
	if text == nil {
		t.Fatal("Expected text to be created")
	}

	if text.Content != "This is a test text content" {
		t.Errorf("Expected content 'This is a test text content', got '%s'", text.Content)
	}

	if text.Name != "Test Text" {
		t.Errorf("Expected name 'Test Text', got '%s'", text.Name)
	}
}

func TestApplyTextEffect(t *testing.T) {
	system := word.NewAdvancedTextSystem()

	// 创建文本
	text := system.CreateAdvancedText("Test text", "This is a test text content")
	if text == nil {
		t.Fatal("Expected text to be created")
	}

	// 应用文本效果
	properties := &word.EffectProperties{
		Color:   "666666",
		Size:    2.0,
		Opacity: 0.5,
		X:       1.0,
		Y:       1.0,
		Blur:    2.0,
	}

	err := system.ApplyTextEffect(text.ID, word.ShadowEffect, properties)
	if err != nil {
		t.Fatalf("Failed to apply text effect: %v", err)
	}

	// 验证效果是否被应用
	if len(text.Effects) == 0 {
		t.Error("Expected text effects to be applied")
	}

	appliedEffect := text.Effects[0]
	if appliedEffect.Type != word.ShadowEffect {
		t.Errorf("Expected effect type ShadowEffect, got %v", appliedEffect.Type)
	}

	if appliedEffect.Properties.Color != "666666" {
		t.Errorf("Expected effect color '666666', got '%s'", appliedEffect.Properties.Color)
	}
}

func TestGetTextSummary(t *testing.T) {
	system := word.NewAdvancedTextSystem()

	// 创建文本
	text := system.CreateAdvancedText("Test Text", "This is a test text with multiple words")
	if text == nil {
		t.Fatal("Expected text to be created")
	}

	// 应用效果
	properties := &word.EffectProperties{
		Color:   "FFFF00",
		Size:    2.0,
		Opacity: 0.5,
	}
	err := system.ApplyTextEffect(text.ID, word.GlowEffect, properties)
	if err != nil {
		t.Fatalf("Failed to apply text effect: %v", err)
	}

	summary := system.GetTextSummary()

	if summary == "" {
		t.Error("Expected non-empty text summary")
	}

	// 检查摘要是否包含预期的文本信息
	expectedInfo := []string{"文本数量", "总段落数", "总运行数"}
	for _, expected := range expectedInfo {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

func TestTextValidation(t *testing.T) {
	system := word.NewAdvancedTextSystem()

	// 创建文本
	text := system.CreateAdvancedText("Test Text", "This is a test text with some spelling errors like 'recieve' and 'seperate'.")
	if text == nil {
		t.Fatal("Expected text to be created")
	}

	// 验证文本 - 这里我们只是测试系统是否正常工作
	// 实际的文本验证功能需要更复杂的实现
	if system.TextValidator == nil {
		t.Error("Expected text validator to be initialized")
	}

	// 检查文本是否被正确创建
	if text.Content == "" {
		t.Error("Expected text content to be set")
	}
}

func TestFormattedTextContent(t *testing.T) {

	// 创建格式化文本内容
	formattedText := word.FormattedTextContent{
		Paragraphs: []word.AdvancedParagraph{
			{
				Text: "This is formatted text",
				Properties: &word.ParagraphProperties{
					Alignment: "left",
				},
				Runs: []word.AdvancedTextRun{
					{
						Text: "This is ",
						Properties: &word.RunProperties{
							Font: &word.RunFont{
								Name:  "Arial",
								Size:  12.0,
								Bold:  false,
								Color: "000000",
							},
						},
					},
					{
						Text: "formatted",
						Properties: &word.RunProperties{
							Font: &word.RunFont{
								Name:  "Arial",
								Size:  14.0,
								Bold:  true,
								Color: "FF0000",
							},
						},
					},
					{
						Text: " text",
						Properties: &word.RunProperties{
							Font: &word.RunFont{
								Name:  "Arial",
								Size:  12.0,
								Bold:  false,
								Color: "000000",
							},
						},
					},
				},
			},
		},
		Language:  "en-US",
		Direction: word.LeftToRight,
	}

	// 验证格式化文本
	if len(formattedText.Paragraphs) != 1 {
		t.Errorf("Expected 1 paragraph, got %d", len(formattedText.Paragraphs))
	}

	paragraph := formattedText.Paragraphs[0]
	if paragraph.Text != "This is formatted text" {
		t.Errorf("Expected content 'This is formatted text', got '%s'", paragraph.Text)
	}

	if len(paragraph.Runs) != 3 {
		t.Errorf("Expected 3 runs, got %d", len(paragraph.Runs))
	}

	// 验证第二个run的格式
	boldRun := paragraph.Runs[1]
	if !boldRun.Properties.Font.Bold {
		t.Error("Expected second run to be bold")
	}

	if boldRun.Properties.Font.Size != 14.0 {
		t.Errorf("Expected second run font size 14.0, got %f", boldRun.Properties.Font.Size)
	}

	if boldRun.Properties.Font.Color != "FF0000" {
		t.Errorf("Expected second run color 'FF0000', got '%s'", boldRun.Properties.Font.Color)
	}
}

func TestAdvancedParagraph(t *testing.T) {

	// 创建高级段落
	paragraph := word.AdvancedParagraph{
		Text: "This is a test paragraph with multiple sentences. It contains various formatting options.",
		Properties: &word.ParagraphProperties{
			Alignment: "justify",
			Spacing:   &word.ParagraphSpacing{Before: 6, After: 6},
		},
		Runs: []word.AdvancedTextRun{
			{
				Text: "This is a test paragraph ",
				Properties: &word.RunProperties{
					Font: &word.RunFont{
						Name: "Arial",
						Size: 12.0,
						Bold: false,
					},
				},
			},
			{
				Text: "with multiple sentences",
				Properties: &word.RunProperties{
					Font: &word.RunFont{
						Name:  "Arial",
						Size:  12.0,
						Bold:  true,
						Color: "0000FF",
					},
				},
			},
		},
	}

	// 验证段落属性
	if paragraph.Properties.Alignment != "justify" {
		t.Errorf("Expected alignment 'justify', got '%s'", paragraph.Properties.Alignment)
	}

	if paragraph.Properties.Spacing.Before != 6 {
		t.Errorf("Expected spacing before 6, got %f", paragraph.Properties.Spacing.Before)
	}

	if paragraph.Properties.Spacing.After != 6 {
		t.Errorf("Expected spacing after 6, got %f", paragraph.Properties.Spacing.After)
	}

	// 验证runs
	if len(paragraph.Runs) != 2 {
		t.Errorf("Expected 2 runs, got %d", len(paragraph.Runs))
	}

	boldRun := paragraph.Runs[1]
	if !boldRun.Properties.Font.Bold {
		t.Error("Expected second run to be bold")
	}

	if boldRun.Properties.Font.Color != "0000FF" {
		t.Errorf("Expected second run color '0000FF', got '%s'", boldRun.Properties.Font.Color)
	}
}

func TestTextEffects(t *testing.T) {
	system := word.NewAdvancedTextSystem()

	// 创建文本效果
	effects := []word.TextEffect{
		{
			ID:   "shadow1",
			Name: "Shadow Effect",
			Type: word.ShadowEffect,
			Properties: &word.EffectProperties{
				Color: "666666",
				X:     2.0,
				Y:     2.0,
				Blur:  3.0,
			},
		},
		{
			ID:   "glow1",
			Name: "Glow Effect",
			Type: word.GlowEffect,
			Properties: &word.EffectProperties{
				Color:   "FFFF00",
				Size:    5.0,
				Opacity: 0.7,
			},
		},
		{
			ID:   "bevel1",
			Name: "Bevel Effect",
			Type: word.BevelEffect,
			Properties: &word.EffectProperties{
				Color: "000000",
				Size:  1.0,
			},
		},
	}

	// 创建文本
	text := system.CreateAdvancedText("Text with effects", "This is a test text with effects")
	if text == nil {
		t.Fatal("Expected text to be created")
	}

	// 应用多个效果
	for _, effect := range effects {
		err := system.ApplyTextEffect(text.ID, effect.Type, effect.Properties)
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
	if shadowEffect.Type != word.ShadowEffect {
		t.Errorf("Expected first effect type ShadowEffect, got %v", shadowEffect.Type)
	}

	if shadowEffect.Properties.X != 2.0 {
		t.Errorf("Expected shadow offset X 2.0, got %f", shadowEffect.Properties.X)
	}
}

func TestTextStatistics(t *testing.T) {
	system := word.NewAdvancedTextSystem()

	// 创建文本
	text := system.CreateAdvancedText("Test Statistics", "This is a test text with multiple words and sentences. It contains various characters and formatting.")
	if text == nil {
		t.Fatal("Expected text to be created")
	}

	// 获取文本统计信息
	stats := system.TextManager.Statistics

	if stats.TotalTexts != 1 {
		t.Errorf("Expected total texts 1, got %d", stats.TotalTexts)
	}

	if stats.TotalWords == 0 {
		t.Error("Expected word count to be greater than 0")
	}

	if stats.TotalCharacters == 0 {
		t.Error("Expected character count to be greater than 0")
	}
}

func TestTextStyles(t *testing.T) {
	system := word.NewAdvancedTextSystem()

	// 创建文本样式
	style := word.AdvancedTextStyle{
		ID:          "TestStyle",
		Name:        "Test Text Style",
		Description: "A test text style",
		Properties: &word.AdvancedTextProperties{
			Language: "en-US",
		},
		BasedOn: "",
		Next:    "",
		Hidden:  false,
		Locked:  false,
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

	if found.Name != "Test Text Style" {
		t.Errorf("Expected style name 'Test Text Style', got '%s'", found.Name)
	}
}

func TestTextValidationRules(t *testing.T) {
	system := word.NewAdvancedTextSystem()

	// 创建验证规则
	rule := word.TextValidationRule{
		ID:          "SpellingCheck",
		Name:        "Spelling Check",
		Description: "Check for spelling errors",
		Type:        1, // SpellingRule
		Condition:   "\\b\\w+\\b",
		Severity:    1, // WarningSeverity
		Enabled:     true,
		Priority:    1,
	}

	// 添加验证规则
	system.TextValidator.Rules = append(system.TextValidator.Rules, rule)

	// 验证规则是否被添加
	if len(system.TextValidator.Rules) == 0 {
		t.Error("Expected validation rules to be added")
	}

	found := false
	for _, r := range system.TextValidator.Rules {
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
