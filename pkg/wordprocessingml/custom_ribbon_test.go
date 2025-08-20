package wordprocessingml

import (
	"testing"
)

// TestNewCustomRibbon 测试创建自定义功能区
func TestNewCustomRibbon(t *testing.T) {
	// 创建默认配置
	config := &RibbonConfig{
		Enabled:           true,
		MaxTabs:           10,
		MaxGroupsPerTab:   5,
		MaxControlsPerGroup: 10,
		Theme:             "default",
		AnimationEnabled:  true,
		TooltipEnabled:    true,
		ShortcutEnabled:   true,
		AutoSave:          true,
		AutoSaveInterval:  300,
	}

	// 创建文档（模拟）
	doc := &Document{}

	// 测试创建自定义功能区
	ribbon := NewCustomRibbon(doc, config)
	if ribbon == nil {
		t.Fatal("自定义功能区创建失败")
	}

	// 验证配置
	if ribbon.config.MaxTabs != 10 {
		t.Errorf("期望最大选项卡数量为10，实际为%d", ribbon.config.MaxTabs)
	}

	if ribbon.config.MaxGroupsPerTab != 5 {
		t.Errorf("期望每个选项卡最大组数为5，实际为%d", ribbon.config.MaxGroupsPerTab)
	}

	if !ribbon.config.Enabled {
		t.Error("功能区应该被启用")
	}
}

// TestCustomRibbonWithNilConfig 测试使用nil配置创建自定义功能区
func TestCustomRibbonWithNilConfig(t *testing.T) {
	// 创建文档（模拟）
	doc := &Document{}

	// 测试使用nil配置创建
	ribbon := NewCustomRibbon(doc, nil)
	if ribbon == nil {
		t.Fatal("使用nil配置创建自定义功能区失败")
	}

	// 验证使用默认配置
	if ribbon.config == nil {
		t.Error("应该使用默认配置")
	}
}

// TestCustomRibbonWithEmptyDocument 测试使用空文档创建自定义功能区
func TestCustomRibbonWithEmptyDocument(t *testing.T) {
	// 创建配置
	config := &RibbonConfig{
		Enabled: true,
		MaxTabs: 5,
	}

	// 测试使用nil文档创建
	ribbon := NewCustomRibbon(nil, config)
	if ribbon == nil {
		t.Fatal("使用nil文档创建自定义功能区失败")
	}

	// 验证文档字段
	if ribbon.document != nil {
		t.Error("文档字段应该为nil")
	}
}
