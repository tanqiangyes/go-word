package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewAdvancedStyleSystem(t *testing.T) {
	system := wordprocessingml.NewAdvancedStyleSystem()
	
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
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	style := wordprocessingml.ParagraphStyleDefinition{
		ID:          "Heading1",
		Name:        "Heading 1",
		BasedOn:     "Normal",
		Next:        "Normal",
		FontSize:    16,
		FontBold:    true,
		Alignment:   "left",
		LineSpacing: 1.5,
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
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	style := wordprocessingml.CharacterStyleDefinition{
		ID:        "Strong",
		Name:      "Strong",
		BasedOn:   "DefaultParagraphFont",
		FontBold:  true,
		FontSize:  12,
		FontColor: "000000",
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
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	style := wordprocessingml.TableStyleDefinition{
		ID:              "TableGrid",
		Name:            "Table Grid",
		BasedOn:         "TableNormal",
		TableBorders:    true,
		HeaderRowFormat: true,
		FirstRowFormat:  true,
		LastRowFormat:   true,
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
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	// 添加段落样式
	paragraphStyle := wordprocessingml.ParagraphStyleDefinition{
		ID:   "TestParagraph",
		Name: "Test Paragraph Style",
	}
	system.AddParagraphStyle(paragraphStyle)
	
	// 添加字符样式
	characterStyle := wordprocessingml.CharacterStyleDefinition{
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
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	// 创建继承链：Normal -> Heading1 -> Heading2
	normalStyle := wordprocessingml.ParagraphStyleDefinition{
		ID:   "Normal",
		Name: "Normal",
	}
	
	heading1Style := wordprocessingml.ParagraphStyleDefinition{
		ID:      "Heading1",
		Name:    "Heading 1",
		BasedOn: "Normal",
	}
	
	heading2Style := wordprocessingml.ParagraphStyleDefinition{
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
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	// 创建样式
	style := wordprocessingml.ParagraphStyleDefinition{
		ID:          "TestStyle",
		Name:        "Test Style",
		FontSize:    14,
		FontBold:    true,
		Alignment:   "center",
		LineSpacing: 1.2,
	}
	
	system.AddParagraphStyle(style)
	
	// 创建测试内容
	content := "Test content for style application"
	
	// 应用样式
	result := system.ApplyStyle("TestStyle", content)
	
	if result == "" {
		t.Error("Expected non-empty result after style application")
	}
	
	// 验证结果包含样式信息
	if !contains(result, "TestStyle") {
		t.Error("Expected result to contain style information")
	}
}

func TestGetStyleSummary(t *testing.T) {
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	// 添加多个样式
	paragraphStyle := wordprocessingml.ParagraphStyleDefinition{
		ID:   "Paragraph1",
		Name: "Test Paragraph",
	}
	
	characterStyle := wordprocessingml.CharacterStyleDefinition{
		ID:   "Character1",
		Name: "Test Character",
	}
	
	tableStyle := wordprocessingml.TableStyleDefinition{
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
	expectedInfo := []string{"Paragraph1", "Character1", "Table1", "Total Styles"}
	for _, expected := range expectedInfo {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

func TestStyleConflictResolution(t *testing.T) {
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	// 创建冲突的样式
	style1 := wordprocessingml.ParagraphStyleDefinition{
		ID:       "ConflictStyle",
		Name:     "Conflict Style 1",
		FontSize: 12,
		FontBold: true,
	}
	
	style2 := wordprocessingml.ParagraphStyleDefinition{
		ID:       "ConflictStyle",
		Name:     "Conflict Style 2",
		FontSize: 14,
		FontBold: false,
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
	
	if found.FontSize != 12 {
		t.Errorf("Expected font size 12, got %d", found.FontSize)
	}
}

// 辅助函数
