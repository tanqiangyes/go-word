package plugin

import (
	"context"
	"testing"
)

// MockPlugin 模拟插件
type MockPlugin struct {
	info   *PluginInfo
	config map[string]interface{}
}

func (mp *MockPlugin) GetInfo() *PluginInfo {
	return mp.info
}

func (mp *MockPlugin) Initialize(config map[string]interface{}) error {
	mp.config = config
	return nil
}

func (mp *MockPlugin) Execute(ctx context.Context, args map[string]interface{}) (*PluginResult, error) {
	return &PluginResult{
		Success: true,
		Data: map[string]interface{}{
			"message": "Mock plugin executed successfully",
		},
		Message: "执行成功",
	}, nil
}

func (mp *MockPlugin) Cleanup() error {
	mp.config = nil
	return nil
}

// TestNewPluginManager 测试创建插件管理器
func TestNewPluginManager(t *testing.T) {
	pm := NewPluginManager()
	if pm == nil {
		t.Fatal("插件管理器创建失败")
	}
	
	if pm.plugins == nil {
		t.Error("插件映射未初始化")
	}
	
	if pm.pluginInfos == nil {
		t.Error("插件信息映射未初始化")
	}
	
	if pm.metrics == nil {
		t.Error("插件指标未初始化")
	}
}

// TestRegisterPlugin 测试注册插件
func TestRegisterPlugin(t *testing.T) {
	pm := NewPluginManager()
	
	// 创建模拟插件
	mockPlugin := &MockPlugin{
		info: &PluginInfo{
			ID:          "test_plugin",
			Name:        "测试插件",
			Version:     "1.0.0",
			Description: "用于测试的插件",
			Category:    "test",
		},
	}
	
	// 注册插件
	err := pm.RegisterPlugin(mockPlugin)
	if err != nil {
		t.Fatalf("插件注册失败: %v", err)
	}
	
	// 验证插件已注册
	if len(pm.plugins) != 1 {
		t.Error("插件数量不匹配")
	}
	
	if pm.metrics.TotalPlugins != 1 {
		t.Error("插件总数指标不匹配")
	}
}

// TestRegisterDuplicatePlugin 测试注册重复插件
func TestRegisterDuplicatePlugin(t *testing.T) {
	pm := NewPluginManager()
	
	// 创建模拟插件
	mockPlugin := &MockPlugin{
		info: &PluginInfo{
			ID:          "test_plugin",
			Name:        "测试插件",
			Version:     "1.0.0",
			Description: "用于测试的插件",
			Category:    "test",
		},
	}
	
	// 第一次注册
	err := pm.RegisterPlugin(mockPlugin)
	if err != nil {
		t.Fatalf("第一次插件注册失败: %v", err)
	}
	
	// 第二次注册相同ID的插件
	err = pm.RegisterPlugin(mockPlugin)
	if err == nil {
		t.Error("应该阻止重复插件注册")
	}
}

// TestUnregisterPlugin 测试注销插件
func TestUnregisterPlugin(t *testing.T) {
	pm := NewPluginManager()
	
	// 创建并注册插件
	mockPlugin := &MockPlugin{
		info: &PluginInfo{
			ID:          "test_plugin",
			Name:        "测试插件",
			Version:     "1.0.0",
			Description: "用于测试的插件",
			Category:    "test",
		},
	}
	
	err := pm.RegisterPlugin(mockPlugin)
	if err != nil {
		t.Fatalf("插件注册失败: %v", err)
	}
	
	// 注销插件
	err = pm.UnregisterPlugin("test_plugin")
	if err != nil {
		t.Fatalf("插件注销失败: %v", err)
	}
	
	// 验证插件已注销
	if len(pm.plugins) != 0 {
		t.Error("插件数量应该为0")
	}
	
	if pm.metrics.TotalPlugins != 0 {
		t.Error("插件总数指标应该为0")
	}
}

// TestConfigurePlugin 测试配置插件
func TestConfigurePlugin(t *testing.T) {
	pm := NewPluginManager()
	
	// 创建并注册插件
	mockPlugin := &MockPlugin{
		info: &PluginInfo{
			ID:          "test_plugin",
			Name:        "测试插件",
			Version:     "1.0.0",
			Description: "用于测试的插件",
			Category:    "test",
		},
	}
	
	err := pm.RegisterPlugin(mockPlugin)
	if err != nil {
		t.Fatalf("插件注册失败: %v", err)
	}
	
	// 配置插件
	config := map[string]interface{}{
		"setting1": "value1",
		"setting2": 42,
	}
	
	err = pm.ConfigurePlugin("test_plugin", config)
	if err != nil {
		t.Fatalf("插件配置失败: %v", err)
	}
	
	// 验证配置已保存
	if pm.metrics.ActivePlugins != 1 {
		t.Error("活跃插件数量不匹配")
	}
}

// TestExecutePlugin 测试执行插件
func TestExecutePlugin(t *testing.T) {
	pm := NewPluginManager()
	
	// 创建并注册插件
	mockPlugin := &MockPlugin{
		info: &PluginInfo{
			ID:          "test_plugin",
			Name:        "测试插件",
			Version:     "1.0.0",
			Description: "用于测试的插件",
			Category:    "test",
		},
	}
	
	err := pm.RegisterPlugin(mockPlugin)
	if err != nil {
		t.Fatalf("插件注册失败: %v", err)
	}
	
	// 配置插件
	config := map[string]interface{}{
		"setting1": "value1",
	}
	
	err = pm.ConfigurePlugin("test_plugin", config)
	if err != nil {
		t.Fatalf("插件配置失败: %v", err)
	}
	
	// 执行插件
	ctx := context.Background()
	args := map[string]interface{}{
		"param1": "value1",
	}
	
	result, err := pm.ExecutePlugin(ctx, "test_plugin", args)
	if err != nil {
		t.Fatalf("插件执行失败: %v", err)
	}
	
	// 验证执行结果
	if !result.Success {
		t.Error("插件执行应该成功")
	}
	
	if result.ExecutionTime == 0 {
		t.Error("执行时间应该大于0")
	}
	
	// 验证指标
	if pm.metrics.TotalExecutions != 1 {
		t.Error("总执行次数不匹配")
	}
	
	if pm.metrics.SuccessCount != 1 {
		t.Error("成功次数不匹配")
	}
}

// TestGetPlugin 测试获取插件
func TestGetPlugin(t *testing.T) {
	pm := NewPluginManager()
	
	// 创建并注册插件
	mockPlugin := &MockPlugin{
		info: &PluginInfo{
			ID:          "test_plugin",
			Name:        "测试插件",
			Version:     "1.0.0",
			Description: "用于测试的插件",
			Category:    "test",
		},
	}
	
	err := pm.RegisterPlugin(mockPlugin)
	if err != nil {
		t.Fatalf("插件注册失败: %v", err)
	}
	
	// 获取插件
	plugin, err := pm.GetPlugin("test_plugin")
	if err != nil {
		t.Fatalf("获取插件失败: %v", err)
	}
	
	if plugin == nil {
		t.Error("插件不应该为nil")
	}
	
	// 获取不存在的插件
	_, err = pm.GetPlugin("nonexistent_plugin")
	if err == nil {
		t.Error("应该返回错误")
	}
}

// TestListPlugins 测试列出插件
func TestListPlugins(t *testing.T) {
	pm := NewPluginManager()
	
	// 创建并注册多个插件
	plugins := []*MockPlugin{
		{
			info: &PluginInfo{
				ID:          "plugin1",
				Name:        "插件1",
				Version:     "1.0.0",
				Category:    "test",
			},
		},
		{
			info: &PluginInfo{
				ID:          "plugin2",
				Name:        "插件2",
				Version:     "1.0.0",
				Category:    "test",
			},
		},
	}
	
	for _, plugin := range plugins {
		err := pm.RegisterPlugin(plugin)
		if err != nil {
			t.Fatalf("插件注册失败: %v", err)
		}
	}
	
	// 列出插件
	pluginList := pm.ListPlugins()
	if len(pluginList) != 2 {
		t.Errorf("插件列表长度不匹配，期望: 2, 实际: %d", len(pluginList))
	}
}

// TestGetMetrics 测试获取指标
func TestGetMetrics(t *testing.T) {
	pm := NewPluginManager()
	
	// 获取初始指标
	metrics := pm.GetMetrics()
	if metrics.TotalPlugins != 0 {
		t.Error("初始插件总数应该为0")
	}
	
	// 注册插件
	mockPlugin := &MockPlugin{
		info: &PluginInfo{
			ID:          "test_plugin",
			Name:        "测试插件",
			Version:     "1.0.0",
			Category:    "test",
		},
	}
	
	err := pm.RegisterPlugin(mockPlugin)
	if err != nil {
		t.Fatalf("插件注册失败: %v", err)
	}
	
	// 获取更新后的指标
	metrics = pm.GetMetrics()
	if metrics.TotalPlugins != 1 {
		t.Error("插件总数应该为1")
	}
}
