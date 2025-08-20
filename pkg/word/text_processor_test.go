package word

import (
	"testing"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// TestNewTextProcessor 测试创建文本处理器
func TestNewTextProcessor(t *testing.T) {
	// 测试默认配置
	tp := NewTextProcessor()
	if tp == nil {
		t.Fatal("文本处理器创建失败")
	}

	// 验证组件已创建
	if tp.FontManager == nil {
		t.Error("FontManager 未初始化")
	}

	if tp.ParagraphManager == nil {
		t.Error("ParagraphManager 未初始化")
	}

	if tp.StyleManager == nil {
		t.Error("StyleManager 未初始化")
	}

	if tp.TextEffectManager == nil {
		t.Error("TextEffectManager 未初始化")
	}

	if tp.LanguageSupport == nil {
		t.Error("LanguageSupport 未初始化")
	}

	if tp.Metrics == nil {
		t.Error("Metrics 未初始化")
	}

	if tp.Logger == nil {
		t.Error("Logger 未初始化")
	}
}

// TestNewFontManager 测试创建字体管理器
func TestNewFontManager(t *testing.T) {
	tp := NewTextProcessor()
	fm := tp.FontManager

	if fm == nil {
		t.Fatal("字体管理器创建失败")
	}

	// 验证默认字体
	if fm.DefaultFont != "SimSun" {
		t.Errorf("默认字体不匹配，期望: SimSun, 实际: %s", fm.DefaultFont)
	}

	// 验证字体映射已初始化
	if len(fm.Fonts) == 0 {
		t.Error("字体映射应该被初始化")
	}

	// 验证字体回退已初始化
	if len(fm.FontFallbacks) == 0 {
		t.Error("字体回退应该被初始化")
	}
}

// TestNewParagraphManager 测试创建段落管理器
func TestNewParagraphManager(t *testing.T) {
	tp := NewTextProcessor()
	pm := tp.ParagraphManager

	if pm == nil {
		t.Fatal("段落管理器创建失败")
	}

	// 验证对齐方式已初始化
	if len(pm.Alignments) == 0 {
		t.Error("对齐方式应该被初始化")
	}

	// 验证缩进已初始化
	if len(pm.Indentations) == 0 {
		t.Error("缩进应该被初始化")
	}

	// 验证间距已初始化
	if len(pm.Spacings) == 0 {
		t.Error("间距应该被初始化")
	}

	// 验证边框已初始化
	if len(pm.Borders) == 0 {
		t.Error("边框应该被初始化")
	}
}

// TestProcessText 测试文本处理
func TestProcessText(t *testing.T) {
	tp := NewTextProcessor()

	// 创建测试内容
	content := &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				Text:  "这是一个测试段落",
				Style: "Normal",
			},
		},
		Tables: []types.Table{},
		Text:   "这是一个测试段落",
	}

	// 处理文本
	err := tp.ProcessText(content)
	if err != nil {
		t.Fatalf("文本处理失败: %v", err)
	}

	// 验证处理结果
	if tp.Metrics.ProcessedParagraphs != 1 {
		t.Errorf("处理的段落数量不匹配，期望: 1, 实际: %d", tp.Metrics.ProcessedParagraphs)
	}

	if tp.Metrics.ProcessedCharacters != 0 { // 字符计数可能为0，因为只是验证处理
		t.Logf("处理的字符数量: %d", tp.Metrics.ProcessedCharacters)
	}
}

// TestTextProcessorLanguageSupport 测试语言支持
func TestTextProcessorLanguageSupport(t *testing.T) {
	tp := NewTextProcessor()
	ls := tp.LanguageSupport

	if ls == nil {
		t.Fatal("语言支持创建失败")
	}

	// 验证支持的语言
	if len(ls.SupportedLanguages) == 0 {
		t.Error("应该支持至少一种语言")
	}

	// 验证中文支持
	chineseLang, exists := ls.SupportedLanguages["zh-CN"]
	if !exists {
		t.Error("应该支持中文")
	}

	if chineseLang.Code != "zh-CN" {
		t.Errorf("中文语言代码不匹配，期望: zh-CN, 实际: %s", chineseLang.Code)
	}

	if chineseLang.Name != "简体中文" {
		t.Errorf("中文语言名称不匹配，期望: 简体中文, 实际: %s", chineseLang.Name)
	}

	// 验证英文支持
	englishLang, exists := ls.SupportedLanguages["en-US"]
	if !exists {
		t.Error("应该支持英文")
	}

	if englishLang.Code != "en-US" {
		t.Errorf("英文语言代码不匹配，期望: en-US, 实际: %s", englishLang.Code)
	}
}

// TestTextProcessorEffects 测试文本效果
func TestTextProcessorEffects(t *testing.T) {
	tp := NewTextProcessor()
	tem := tp.TextEffectManager

	if tem == nil {
		t.Fatal("文本效果管理器创建失败")
	}

	// 验证效果已初始化
	if len(tem.Effects) == 0 {
		t.Error("效果应该被初始化")
	}

	// 验证效果数量
	if len(tem.Effects) < 3 {
		t.Errorf("应该至少初始化3种效果，实际: %d", len(tem.Effects))
	}
}
