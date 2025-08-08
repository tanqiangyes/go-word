package tests

import (
	"testing"
	"time"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
	"github.com/tanqiangyes/go-word/pkg/utils"
)

// TestNewAdvancedDocumentProcessor 测试创建高级文档处理器
func TestNewAdvancedDocumentProcessor(t *testing.T) {
	adp := wordprocessingml.NewAdvancedDocumentProcessor()
	
	if adp == nil {
		t.Fatal("高级文档处理器创建失败")
	}
	
	// 检查组件是否正确初始化
	if adp.GetTextProcessor() == nil {
		t.Error("文字处理器未正确初始化")
	}
	
	if adp.GetLayoutManager() == nil {
		t.Error("排版管理器未正确初始化")
	}
	
	if adp.GetThemeManager() == nil {
		t.Error("主题管理器未正确初始化")
	}
	
	if adp.GetStyleLibrary() == nil {
		t.Error("样式库未正确初始化")
	}
	
	if adp.GetFormatOptimizer() == nil {
		t.Error("格式优化器未正确初始化")
	}
}

// TestTextProcessor 测试文字处理器
func TestTextProcessor(t *testing.T) {
	tp := wordprocessingml.NewTextProcessor()
	
	if tp == nil {
		t.Fatal("文字处理器创建失败")
	}
	
	// 创建测试文档内容
	content := &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				Runs: []types.Run{
					{Text: "这是一个测试段落"},
					{Text: "包含多个文本运行"},
				},
			},
			{
				Runs: []types.Run{
					{Text: "第二个段落"},
				},
			},
		},
	}
	
	// 测试文字处理
	if err := tp.ProcessText(content); err != nil {
		t.Errorf("文字处理失败: %v", err)
	}
	
	// 检查性能指标
	metrics := tp.GetMetrics()
	if metrics.ProcessedParagraphs != 2 {
		t.Errorf("期望处理 2 个段落，实际处理了 %d 个", metrics.ProcessedParagraphs)
	}
	
	if metrics.ProcessedCharacters == 0 {
		t.Error("字符处理计数应该大于 0")
	}
}

// TestLayoutManager 测试排版管理器
func TestLayoutManager(t *testing.T) {
	lm := wordprocessingml.NewLayoutManager()
	
	if lm == nil {
		t.Fatal("排版管理器创建失败")
	}
	
	// 创建测试文档内容
	content := &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				Runs: []types.Run{
					{Text: "测试段落布局"},
				},
			},
		},
		Tables: []types.Table{
			{
				Rows: []types.TableRow{
					{
						Cells: []types.TableCell{
							{Text: "测试表格布局"},
						},
					},
				},
			},
		},
	}
	
	// 测试布局处理
	if err := lm.ProcessLayout(content); err != nil {
		t.Errorf("布局处理失败: %v", err)
	}
	
	// 检查性能指标
	metrics := lm.GetMetrics()
	if metrics.ElementsPositioned != 2 {
		t.Errorf("期望定位 2 个元素，实际定位了 %d 个", metrics.ElementsPositioned)
	}
}

// TestThemeManager 测试主题管理器
func TestThemeManager(t *testing.T) {
	tm := wordprocessingml.NewThemeManager()
	
	if tm == nil {
		t.Fatal("主题管理器创建失败")
	}
	
	// 测试获取当前主题
	currentTheme := tm.GetCurrentTheme()
	if currentTheme == nil {
		t.Error("当前主题不应该为 nil")
	}
	
	if currentTheme.ID != "default" {
		t.Errorf("期望默认主题ID为 'default'，实际为 '%s'", currentTheme.ID)
	}
	
	// 测试列出所有主题
	themes := tm.ListThemes()
	if len(themes) == 0 {
		t.Error("主题列表不应该为空")
	}
	
	// 测试创建新主题
	newTheme := tm.CreateTheme("测试主题", "这是一个测试主题")
	if newTheme == nil {
		t.Error("新主题创建失败")
	}
	
	if newTheme.Name != "测试主题" {
		t.Errorf("期望主题名称为 '测试主题'，实际为 '%s'", newTheme.Name)
	}
	
	// 测试获取指定主题
	theme, err := tm.GetTheme(newTheme.ID)
	if err != nil {
		t.Errorf("获取主题失败: %v", err)
	}
	
	if theme.ID != newTheme.ID {
		t.Errorf("期望主题ID为 '%s'，实际为 '%s'", newTheme.ID, theme.ID)
	}
	
	// 测试创建颜色方案
	colorScheme := tm.CreateColorScheme("测试颜色方案", "这是一个测试颜色方案")
	if colorScheme == nil {
		t.Error("颜色方案创建失败")
	}
	
	// 测试创建字体方案
	fontScheme := tm.CreateFontScheme("测试字体方案", "这是一个测试字体方案")
	if fontScheme == nil {
		t.Error("字体方案创建失败")
	}
}

// TestStyleLibrary 测试样式库
func TestStyleLibrary(t *testing.T) {
	sl := wordprocessingml.NewStyleLibrary()
	
	if sl == nil {
		t.Fatal("样式库创建失败")
	}
	
	// 检查样式库是否正确初始化
	if sl == nil {
		t.Error("样式库不应该为 nil")
	}
}

// TestFormatOptimizer 测试格式优化器
func TestFormatOptimizer(t *testing.T) {
	fo := wordprocessingml.NewFormatOptimizer()
	
	if fo == nil {
		t.Fatal("格式优化器创建失败")
	}
	
	// 检查格式优化器是否正确初始化
	if fo == nil {
		t.Error("格式优化器不应该为 nil")
	}
}

// TestAdvancedDocumentProcessorIntegration 测试高级文档处理器集成
func TestAdvancedDocumentProcessorIntegration(t *testing.T) {
	adp := wordprocessingml.NewAdvancedDocumentProcessor()
	
	// 创建测试文档内容
	content := &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				Runs: []types.Run{
					{Text: "这是第一个段落，用于测试高级文档处理功能。"},
					{Text: "包含多个文本运行，测试文字处理能力。"},
				},
			},
			{
				Runs: []types.Run{
					{Text: "这是第二个段落，测试排版和布局功能。"},
				},
			},
		},
		Tables: []types.Table{
			{
				Rows: []types.TableRow{
					{
						Cells: []types.TableCell{
							{Text: "表格单元格 1"},
							{Text: "表格单元格 2"},
						},
					},
					{
						Cells: []types.TableCell{
							{Text: "表格单元格 3"},
							{Text: "表格单元格 4"},
						},
					},
				},
			},
		},
	}
	
	// 测试完整的文档处理流程
	startTime := time.Now()
	if err := adp.ProcessDocument(content); err != nil {
		t.Errorf("高级文档处理失败: %v", err)
	}
	processingTime := time.Since(startTime)
	
	// 检查处理时间是否合理
	if processingTime > 5*time.Second {
		t.Errorf("处理时间过长: %v", processingTime)
	}
	
	// 检查性能指标
	metrics := adp.GetMetrics()
	if metrics.DocumentsProcessed != 1 {
		t.Errorf("期望处理 1 个文档，实际处理了 %d 个", metrics.DocumentsProcessed)
	}
	
	if metrics.TotalProcessingTime == 0 {
		t.Error("总处理时间应该大于 0")
	}
	
	// 检查各个组件的处理时间
	if metrics.TextProcessingTime == 0 {
		t.Error("文字处理时间应该大于 0")
	}
	
	if metrics.LayoutProcessingTime == 0 {
		t.Error("排版处理时间应该大于 0")
	}
	
	if metrics.ThemeProcessingTime == 0 {
		t.Error("主题处理时间应该大于 0")
	}
	
	if metrics.StyleProcessingTime == 0 {
		t.Error("样式处理时间应该大于 0")
	}
	
	if metrics.OptimizationTime == 0 {
		t.Error("优化时间应该大于 0")
	}
}

// TestTextProcessorMetrics 测试文字处理器性能指标
func TestTextProcessorMetrics(t *testing.T) {
	tp := wordprocessingml.NewTextProcessor()
	
	// 创建大量文本内容进行测试
	content := &types.DocumentContent{
		Paragraphs: make([]types.Paragraph, 100),
	}
	
	for i := range content.Paragraphs {
		content.Paragraphs[i] = types.Paragraph{
			Runs: []types.Run{
				{Text: "这是一个测试段落，用于测试文字处理器的性能。"},
				{Text: "包含多个文本运行，模拟真实的文档内容。"},
			},
		}
	}
	
	// 测试处理性能
	startTime := time.Now()
	if err := tp.ProcessText(content); err != nil {
		t.Errorf("文字处理失败: %v", err)
	}
	processingTime := time.Since(startTime)
	
	// 检查性能指标
	metrics := tp.GetMetrics()
	if metrics.ProcessedParagraphs != 100 {
		t.Errorf("期望处理 100 个段落，实际处理了 %d 个", metrics.ProcessedParagraphs)
	}
	
	if metrics.ProcessedCharacters == 0 {
		t.Error("字符处理计数应该大于 0")
	}
	
	if metrics.ProcessingTime == 0 {
		t.Error("处理时间应该大于 0")
	}
	
	// 检查处理时间是否合理（100个段落应该在合理时间内完成）
	if processingTime > 1*time.Second {
		t.Errorf("处理时间过长: %v", processingTime)
	}
}

// TestLayoutManagerMetrics 测试排版管理器性能指标
func TestLayoutManagerMetrics(t *testing.T) {
	lm := wordprocessingml.NewLayoutManager()
	
	// 创建复杂的布局内容
	content := &types.DocumentContent{
		Paragraphs: make([]types.Paragraph, 50),
		Tables:     make([]types.Table, 10),
	}
	
	// 填充段落
	for i := range content.Paragraphs {
		content.Paragraphs[i] = types.Paragraph{
			Runs: []types.Run{
				{Text: "测试段落布局"},
			},
		}
	}
	
	// 填充表格
	for i := range content.Tables {
		content.Tables[i] = types.Table{
			Rows: []types.TableRow{
				{
					Cells: []types.TableCell{
						{Text: "表格单元格"},
						{Text: "表格单元格"},
					},
				},
				{
					Cells: []types.TableCell{
						{Text: "表格单元格"},
						{Text: "表格单元格"},
					},
				},
			},
		}
	}
	
	// 测试布局处理性能
	startTime := time.Now()
	if err := lm.ProcessLayout(content); err != nil {
		t.Errorf("布局处理失败: %v", err)
	}
	processingTime := time.Since(startTime)
	
	// 检查性能指标
	metrics := lm.GetMetrics()
	expectedElements := 60 // 50个段落 + 10个表格
	if metrics.ElementsPositioned != int64(expectedElements) {
		t.Errorf("期望定位 %d 个元素，实际定位了 %d 个", expectedElements, metrics.ElementsPositioned)
	}
	
	if metrics.ProcessingTime == 0 {
		t.Error("处理时间应该大于 0")
	}
	
	// 检查处理时间是否合理
	if processingTime > 2*time.Second {
		t.Errorf("处理时间过长: %v", processingTime)
	}
}

// TestThemeManagerPerformance 测试主题管理器性能
func TestThemeManagerPerformance(t *testing.T) {
	tm := wordprocessingml.NewThemeManager()
	
	// 创建大量文档内容
	content := &types.DocumentContent{
		Paragraphs: make([]types.Paragraph, 200),
	}
	
	for i := range content.Paragraphs {
		content.Paragraphs[i] = types.Paragraph{
			Runs: []types.Run{
				{Text: "测试主题应用性能"},
			},
		}
	}
	
	// 测试主题应用性能
	startTime := time.Now()
	if err := tm.ApplyTheme("default", content); err != nil {
		t.Errorf("主题应用失败: %v", err)
	}
	processingTime := time.Since(startTime)
	
	// 检查性能指标
	metrics := tm.GetMetrics()
	if metrics.ThemesApplied != 1 {
		t.Errorf("期望应用 1 个主题，实际应用了 %d 个", metrics.ThemesApplied)
	}
	
	if metrics.ProcessingTime == 0 {
		t.Error("处理时间应该大于 0")
	}
	
	// 检查处理时间是否合理
	if processingTime > 1*time.Second {
		t.Errorf("处理时间过长: %v", processingTime)
	}
}

// TestAdvancedDocumentErrorHandling 测试错误处理
func TestAdvancedDocumentErrorHandling(t *testing.T) {
	adp := wordprocessingml.NewAdvancedDocumentProcessor()
	
	// 测试空内容处理
	emptyContent := &types.DocumentContent{}
	if err := adp.ProcessDocument(emptyContent); err != nil {
		t.Errorf("空内容处理失败: %v", err)
	}
	
	// 测试主题管理器错误处理
	tm := adp.GetThemeManager()
	
	// 测试获取不存在的主题
	_, err := tm.GetTheme("nonexistent")
	if err == nil {
		t.Error("获取不存在的主题应该返回错误")
	}
	
	// 测试删除默认主题
	err = tm.DeleteTheme("default")
	if err == nil {
		t.Error("删除默认主题应该返回错误")
	}
}

// TestLoggerIntegration 测试日志器集成
func TestLoggerIntegration(t *testing.T) {
	// 创建自定义日志器
	logger := utils.NewLogger(utils.LogLevelDebug, nil)
	adp := wordprocessingml.NewAdvancedDocumentProcessor()
	
	// 设置自定义日志器
	adp.SetLogger(logger)
	
	// 创建测试内容
	content := &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				Runs: []types.Run{
					{Text: "测试日志器集成"},
				},
			},
		},
	}
	
	// 测试处理（应该产生日志输出）
	if err := adp.ProcessDocument(content); err != nil {
		t.Errorf("文档处理失败: %v", err)
	}
	
	// 检查性能指标
	metrics := adp.GetMetrics()
	if metrics.DocumentsProcessed != 1 {
		t.Errorf("期望处理 1 个文档，实际处理了 %d 个", metrics.DocumentsProcessed)
	}
}
