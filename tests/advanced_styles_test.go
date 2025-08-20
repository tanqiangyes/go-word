package tests

import (
	"testing"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/word"
)

func TestNewAdvancedStyleSystem(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	if system == nil {
		t.Fatal("Expected AdvancedStyleSystem to be created")
	}

	if system.StyleManager == nil {
		t.Error("Expected StyleManager to be initialized")
	}

	if system.StyleCache == nil {
		t.Error("Expected StyleCache to be initialized")
	}

	if system.InheritanceChain == nil {
		t.Error("Expected InheritanceChain to be initialized")
	}

	if system.ConflictResolver == nil {
		t.Error("Expected ConflictResolver to be initialized")
	}
}

func TestAddParagraphStyle(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	style := &word.ParagraphStyleDefinition{
		ID:      "Heading1",
		Name:    "Heading 1",
		BasedOn: "Normal",
		Next:    "Normal",
		Properties: &word.ParagraphStyleProperties{
			Alignment: "left",
		},
	}

	err := system.AddParagraphStyle(style)
	if err != nil {
		t.Fatalf("Failed to add paragraph style: %v", err)
	}

	// 验证样式是否被添加
	found := system.GetParagraphStyle("Heading1")
	if found == nil {
		t.Error("Expected to find added paragraph style")
	}

	if found.ID != "Heading1" {
		t.Errorf("Expected style ID 'Heading1', got '%s'", found.ID)
	}
}

func TestAddCharacterStyle(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	style := &word.CharacterStyleDefinition{
		ID:      "Strong",
		Name:    "Strong",
		BasedOn: "DefaultParagraphFont",
		Properties: &word.CharacterStyleProperties{
			Font: &word.Font{
				Name: "Arial",
			},
		},
	}

	err := system.AddCharacterStyle(style)
	if err != nil {
		t.Fatalf("Failed to add character style: %v", err)
	}

	// 验证样式是否被添加
	found := system.GetCharacterStyle("Strong")
	if found == nil {
		t.Error("Expected to find added character style")
	}

	if found.ID != "Strong" {
		t.Errorf("Expected style ID 'Strong', got '%s'", found.ID)
	}
}

func TestAddTableStyle(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	style := &word.TableStyleDefinition{
		ID:      "TableGrid",
		Name:    "Table Grid",
		BasedOn: "TableNormal",
		Properties: &word.TableStyleProperties{
			Alignment: "left",
		},
	}

	err := system.AddTableStyle(style)
	if err != nil {
		t.Fatalf("Failed to add table style: %v", err)
	}

	// 验证样式是否被添加
	found := system.GetTableStyle("TableGrid")
	if found == nil {
		t.Error("Expected to find added table style")
	}

	if found.ID != "TableGrid" {
		t.Errorf("Expected style ID 'TableGrid', got '%s'", found.ID)
	}
}

func TestGetStyle(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	// 添加段落样式
	paragraphStyle := &word.ParagraphStyleDefinition{
		ID:   "TestParagraph",
		Name: "Test Paragraph Style",
	}
	system.AddParagraphStyle(paragraphStyle)

	// 添加字符样式
	characterStyle := &word.CharacterStyleDefinition{
		ID:   "TestCharacter",
		Name: "Test Character Style",
	}
	system.AddCharacterStyle(characterStyle)

	// 测试获取段落样式
	found := system.GetStyle("TestParagraph")
	if found == nil {
		t.Error("Expected to find paragraph style")
	}

	// 测试获取字符样式
	found = system.GetStyle("TestCharacter")
	if found == nil {
		t.Error("Expected to find character style")
	}

	// 测试获取不存在的样式
	found = system.GetStyle("NonExistent")
	if found != nil {
		t.Error("Expected nil for non-existent style")
	}
}

func TestGetInheritanceChain(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	// 创建继承链：Normal -> Heading1 -> Heading2
	normalStyle := &word.ParagraphStyleDefinition{
		ID:   "Normal",
		Name: "Normal",
	}

	heading1Style := &word.ParagraphStyleDefinition{
		ID:      "Heading1",
		Name:    "Heading 1",
		BasedOn: "Normal",
	}

	heading2Style := &word.ParagraphStyleDefinition{
		ID:      "Heading2",
		Name:    "Heading 2",
		BasedOn: "Heading1",
	}

	system.AddParagraphStyle(normalStyle)
	system.AddParagraphStyle(heading1Style)
	system.AddParagraphStyle(heading2Style)

	// 获取继承链
	chain := system.GetInheritanceChain("Heading2")

	if len(chain) != 3 {
		t.Errorf("Expected inheritance chain length 3, got %d", len(chain))
	}

	// 验证继承链顺序
	expectedChain := []string{"Heading2", "Heading1", "Normal"}
	for i, expected := range expectedChain {
		if i >= len(chain) || chain[i] != expected {
			t.Errorf("Expected chain[%d] to be '%s', got '%s'", i, expected, chain[i])
		}
	}
}

func TestApplyStyle(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	// 创建样式
	style := &word.ParagraphStyleDefinition{
		ID:   "TestStyle",
		Name: "Test Style",
		Properties: &word.ParagraphStyleProperties{
			Alignment: "center",
		},
	}

	system.AddParagraphStyle(style)

	// 创建测试段落
	paragraph := &types.Paragraph{
		Text:  "Test content for style application",
		Style: "",
		Runs: []types.Run{
			{
				Text:     "Test content for style application",
				FontSize: 11,
				FontName: "Arial",
			},
		},
	}

	// 应用样式
	err := system.ApplyStyle(paragraph, "TestStyle")

	if err != nil {
		t.Errorf("Failed to apply style: %v", err)
	}

	// 验证样式是否被应用
	if paragraph.Style != "TestStyle" {
		t.Errorf("Expected paragraph style to be 'TestStyle', got '%s'", paragraph.Style)
	}
}

func TestGetStyleSummary(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	// 添加多个样式
	paragraphStyle := &word.ParagraphStyleDefinition{
		ID:   "Paragraph1",
		Name: "Test Paragraph",
	}

	characterStyle := &word.CharacterStyleDefinition{
		ID:   "Character1",
		Name: "Test Character",
	}

	tableStyle := &word.TableStyleDefinition{
		ID:   "Table1",
		Name: "Test Table",
	}

	system.AddParagraphStyle(paragraphStyle)
	system.AddCharacterStyle(characterStyle)
	system.AddTableStyle(tableStyle)

	summary := system.GetStyleSummary()

	if summary == "" {
		t.Error("Expected non-empty style summary")
	}

	// 检查摘要是否包含预期的样式信息
	expectedInfo := []string{"段落样式", "字符样式", "表格样式", "样式缓存", "继承链", "冲突数量"}
	for _, expected := range expectedInfo {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

func TestStyleConflictResolution(t *testing.T) {
	system := word.NewAdvancedStyleSystem()

	// 创建冲突的样式
	style1 := &word.ParagraphStyleDefinition{
		ID:   "ConflictStyle",
		Name: "Conflict Style 1",
		Properties: &word.ParagraphStyleProperties{
			Alignment: "left",
		},
	}

	style2 := &word.ParagraphStyleDefinition{
		ID:   "ConflictStyle",
		Name: "Conflict Style 2",
		Properties: &word.ParagraphStyleProperties{
			Alignment: "right",
		},
	}

	// 添加第一个样式
	err1 := system.AddParagraphStyle(style1)
	if err1 != nil {
		t.Fatalf("Failed to add first style: %v", err1)
	}

	// 尝试添加冲突的样式
	err2 := system.AddParagraphStyle(style2)
	if err2 == nil {
		t.Error("Expected error when adding conflicting style")
	}

	// 验证原始样式保持不变
	found := system.GetParagraphStyle("ConflictStyle")
	if found == nil {
		t.Error("Expected original style to remain")
	}

	if found.Name != "Conflict Style 1" {
		t.Errorf("Expected style name 'Conflict Style 1', got '%s'", found.Name)
	}
}

// 辅助函数
