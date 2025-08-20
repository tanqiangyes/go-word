package examples

import (
	"context"
	"testing"
)

// TestNewTextFormatterPlugin 测试创建文本格式化插件
func TestNewTextFormatterPlugin(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	if plugin == nil {
		t.Fatal("文本格式化插件创建失败")
	}
}

// TestTextFormatterPlugin_GetInfo 测试获取插件信息
func TestTextFormatterPlugin_GetInfo(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	info := plugin.GetInfo()
	
	if info == nil {
		t.Fatal("插件信息不能为空")
	}
	
	if info.ID != "text_formatter" {
		t.Errorf("插件ID不匹配，期望: text_formatter, 实际: %s", info.ID)
	}
	
	if info.Name != "文本格式化插件" {
		t.Errorf("插件名称不匹配，期望: 文本格式化插件, 实际: %s", info.Name)
	}
	
	if info.Category != "text" {
		t.Errorf("插件类别不匹配，期望: text, 实际: %s", info.Category)
	}
	
	if len(info.Tags) == 0 {
		t.Error("插件标签不能为空")
	}
}

// TestTextFormatterPlugin_Initialize 测试插件初始化
func TestTextFormatterPlugin_Initialize(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	config := map[string]interface{}{
		"setting1": "value1",
		"setting2": 42,
	}
	
	err := plugin.Initialize(config)
	if err != nil {
		t.Fatalf("插件初始化失败: %v", err)
	}
	
	if plugin.config == nil {
		t.Error("插件配置未保存")
	}
}

// TestTextFormatterPlugin_Execute 测试插件执行
func TestTextFormatterPlugin_Execute(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	// 测试大写转换
	ctx := context.Background()
	args := map[string]interface{}{
		"text":   "hello world",
		"format": "uppercase",
	}
	
	result, err := plugin.Execute(ctx, args)
	if err != nil {
		t.Fatalf("插件执行失败: %v", err)
	}
	
	if !result.Success {
		t.Error("插件执行应该成功")
	}
	
	formattedText, ok := result.Data["formatted_text"].(string)
	if !ok {
		t.Fatal("格式化文本不存在")
	}
	
	if formattedText != "HELLO WORLD" {
		t.Errorf("大写转换失败，期望: HELLO WORLD, 实际: %s", formattedText)
	}
}

// TestTextFormatterPlugin_ExecuteLowercase 测试小写转换
func TestTextFormatterPlugin_ExecuteLowercase(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	ctx := context.Background()
	args := map[string]interface{}{
		"text":   "HELLO WORLD",
		"format": "lowercase",
	}
	
	result, err := plugin.Execute(ctx, args)
	if err != nil {
		t.Fatalf("插件执行失败: %v", err)
	}
	
	formattedText, ok := result.Data["formatted_text"].(string)
	if !ok {
		t.Fatal("格式化文本不存在")
	}
	
	if formattedText != "hello world" {
		t.Errorf("小写转换失败，期望: hello world, 实际: %s", formattedText)
	}
}

// TestTextFormatterPlugin_ExecuteTitlecase 测试标题格式转换
func TestTextFormatterPlugin_ExecuteTitlecase(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	ctx := context.Background()
	args := map[string]interface{}{
		"text":   "hello world",
		"format": "titlecase",
	}
	
	result, err := plugin.Execute(ctx, args)
	if err != nil {
		t.Fatalf("插件执行失败: %v", err)
	}
	
	formattedText, ok := result.Data["formatted_text"].(string)
	if !ok {
		t.Fatal("格式化文本不存在")
	}
	
	if formattedText != "Hello World" {
		t.Errorf("标题格式转换失败，期望: Hello World, 实际: %s", formattedText)
	}
}

// TestTextFormatterPlugin_ExecuteTrim 测试去除空格
func TestTextFormatterPlugin_ExecuteTrim(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	ctx := context.Background()
	args := map[string]interface{}{
		"text":   "  hello world  ",
		"format": "trim",
	}
	
	result, err := plugin.Execute(ctx, args)
	if err != nil {
		t.Fatalf("插件执行失败: %v", err)
	}
	
	formattedText, ok := result.Data["formatted_text"].(string)
	if !ok {
		t.Fatal("格式化文本不存在")
	}
	
	if formattedText != "hello world" {
		t.Errorf("去除空格失败，期望: hello world, 实际: %s", formattedText)
	}
}

// TestTextFormatterPlugin_ExecuteNormalize 测试文本标准化
func TestTextFormatterPlugin_ExecuteNormalize(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	ctx := context.Background()
	args := map[string]interface{}{
		"text":   "  hello   world  ,  how   are   you  ?  ",
		"format": "normalize",
	}
	
	result, err := plugin.Execute(ctx, args)
	if err != nil {
		t.Fatalf("插件执行失败: %v", err)
	}
	
	formattedText, ok := result.Data["formatted_text"].(string)
	if !ok {
		t.Fatal("格式化文本不存在")
	}
	
	expected := "hello world, how are you?"
	if formattedText != expected {
		t.Errorf("文本标准化失败，期望: %s, 实际: %s", expected, formattedText)
	}
}

// TestTextFormatterPlugin_ExecuteDefaultFormat 测试默认格式
func TestTextFormatterPlugin_ExecuteDefaultFormat(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	ctx := context.Background()
	args := map[string]interface{}{
		"text": "  hello   world  ",
		// 不指定format，使用默认的normalize
	}
	
	result, err := plugin.Execute(ctx, args)
	if err != nil {
		t.Fatalf("插件执行失败: %v", err)
	}
	
	formattedText, ok := result.Data["formatted_text"].(string)
	if !ok {
		t.Fatal("格式化文本不存在")
	}
	
	if formattedText != "hello world" {
		t.Errorf("默认格式处理失败，期望: hello world, 实际: %s", formattedText)
	}
}

// TestTextFormatterPlugin_ExecuteInvalidFormat 测试无效格式
func TestTextFormatterPlugin_ExecuteInvalidFormat(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	ctx := context.Background()
	args := map[string]interface{}{
		"text":   "hello world",
		"format": "invalid_format",
	}
	
	result, err := plugin.Execute(ctx, args)
	if err == nil {
		t.Error("应该返回错误")
	}
	
	if result.Success {
		t.Error("执行结果应该失败")
	}
}

// TestTextFormatterPlugin_ExecuteMissingText 测试缺少文本参数
func TestTextFormatterPlugin_ExecuteMissingText(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	ctx := context.Background()
	args := map[string]interface{}{
		"format": "uppercase",
		// 缺少text参数
	}
	
	result, err := plugin.Execute(ctx, args)
	if err == nil {
		t.Error("应该返回错误")
	}
	
	if result.Success {
		t.Error("执行结果应该失败")
	}
}

// TestTextFormatterPlugin_Cleanup 测试插件清理
func TestTextFormatterPlugin_Cleanup(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	// 先初始化
	config := map[string]interface{}{
		"setting1": "value1",
	}
	
	err := plugin.Initialize(config)
	if err != nil {
		t.Fatalf("插件初始化失败: %v", err)
	}
	
	// 清理
	err = plugin.Cleanup()
	if err != nil {
		t.Fatalf("插件清理失败: %v", err)
	}
	
	if plugin.config != nil {
		t.Error("插件配置应该被清理")
	}
}

// TestTextFormatterPlugin_normalizeText 测试文本标准化函数
func TestTextFormatterPlugin_normalizeText(t *testing.T) {
	plugin := NewTextFormatterPlugin()
	
	testCases := []struct {
		input    string
		expected string
	}{
		{"  hello   world  ", "hello world"},
		{"  hello,world  ", "hello,world"},
		{"  hello.world  ", "hello.world"},
		{"  hello!world  ", "hello!world"},
		{"  hello?world  ", "hello?world"},
		{"  hello   ,   world   .   ", "hello, world."},
	}
	
	for _, tc := range testCases {
		result := plugin.normalizeText(tc.input)
		if result != tc.expected {
			t.Errorf("文本标准化失败，输入: %q, 期望: %q, 实际: %q", tc.input, tc.expected, result)
		}
	}
}
