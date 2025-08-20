package plugin

import (
	"context"
	"fmt"
	"sync"
	"time"

	"github.com/tanqiangyes/go-word/pkg/utils"
)

// Plugin 插件接口
type Plugin interface {
	// GetInfo 获取插件信息
	GetInfo() *PluginInfo

	// Initialize 初始化插件
	Initialize(config map[string]interface{}) error

	// Execute 执行插件
	Execute(ctx context.Context, args map[string]interface{}) (*PluginResult, error)

	// Cleanup 清理插件资源
	Cleanup() error
}

// PluginInfo 插件信息
type PluginInfo struct {
	ID          string                 `json:"id"`
	Name        string                 `json:"name"`
	Version     string                 `json:"version"`
	Description string                 `json:"description"`
	Author      string                 `json:"author"`
	License     string                 `json:"license"`
	Category    string                 `json:"category"`
	Tags        []string               `json:"tags"`
	Required    []string               `json:"required"`
	Optional    []string               `json:"optional"`
	CreatedAt   time.Time              `json:"created_at"`
	UpdatedAt   time.Time              `json:"updated_at"`
	Metadata    map[string]interface{} `json:"metadata"`
}

// PluginResult 插件执行结果
type PluginResult struct {
	Success       bool                   `json:"success"`
	Data          map[string]interface{} `json:"data"`
	Message       string                 `json:"message"`
	Error         error                  `json:"error,omitempty"`
	ExecutionTime time.Duration          `json:"execution_time"`
	Timestamp     time.Time              `json:"timestamp"`
}

// PluginManager 插件管理器
type PluginManager struct {
	Plugins     map[string]Plugin
	PluginInfos map[string]*PluginInfo
	Configs     map[string]map[string]interface{}
	Results     map[string]*PluginResult
	Mu          sync.RWMutex
	Logger      *utils.Logger
	Metrics     *PluginMetrics
}

// PluginMetrics 插件指标
type PluginMetrics struct {
	TotalPlugins    int64         `json:"total_plugins"`
	ActivePlugins   int64         `json:"active_plugins"`
	TotalExecutions int64         `json:"total_executions"`
	SuccessCount    int64         `json:"success_count"`
	ErrorCount      int64         `json:"error_count"`
	AverageTime     time.Duration `json:"average_time"`
	LastExecution   time.Time     `json:"last_execution"`
}

// NewPluginManager 创建插件管理器
func NewPluginManager() *PluginManager {
	return &PluginManager{
		Plugins:     make(map[string]Plugin),
		PluginInfos: make(map[string]*PluginInfo),
		Configs:     make(map[string]map[string]interface{}),
		Results:     make(map[string]*PluginResult),
		Logger:      utils.NewLogger(utils.LogLevelInfo, nil),
		Metrics:     &PluginMetrics{},
	}
}

// RegisterPlugin 注册插件
func (pm *PluginManager) RegisterPlugin(plugin Plugin) error {
	pm.Mu.Lock()
	defer pm.Mu.Unlock()

	info := plugin.GetInfo()
	if info == nil {
		return fmt.Errorf("插件信息不能为空")
	}

	if info.ID == "" {
		return fmt.Errorf("插件ID不能为空")
	}

	if _, exists := pm.Plugins[info.ID]; exists {
		return fmt.Errorf("插件 %s 已存在", info.ID)
	}

	// 检查依赖
	if err := pm.checkDependencies(info); err != nil {
		return fmt.Errorf("依赖检查失败: %w", err)
	}

	// 注册插件
	pm.Plugins[info.ID] = plugin
	pm.PluginInfos[info.ID] = info
	pm.Metrics.TotalPlugins++

	pm.Logger.Info("插件注册成功，插件ID: %s, 名称: %s, 版本: %s, 类别: %s", info.ID, info.Name, info.Version, info.Category)

	return nil
}

// UnregisterPlugin 注销插件
func (pm *PluginManager) UnregisterPlugin(pluginID string) error {
	pm.Mu.Lock()
	defer pm.Mu.Unlock()

	plugin, exists := pm.Plugins[pluginID]
	if !exists {
		return fmt.Errorf("插件 %s 不存在", pluginID)
	}

	// 清理插件资源
	if err := plugin.Cleanup(); err != nil {
		pm.Logger.Warning("插件清理失败，插件ID: %s, 错误: %s", pluginID, err.Error())
	}

	// 注销插件
	delete(pm.Plugins, pluginID)
	delete(pm.PluginInfos, pluginID)
	delete(pm.Configs, pluginID)
	delete(pm.Results, pluginID)

	pm.Metrics.TotalPlugins--

	pm.Logger.Info("插件注销成功，插件ID: %s", pluginID)

	return nil
}

// ConfigurePlugin 配置插件
func (pm *PluginManager) ConfigurePlugin(pluginID string, config map[string]interface{}) error {
	pm.Mu.Lock()
	defer pm.Mu.Unlock()

	plugin, exists := pm.Plugins[pluginID]
	if !exists {
		return fmt.Errorf("插件 %s 不存在", pluginID)
	}

	// 初始化插件
	if err := plugin.Initialize(config); err != nil {
		return fmt.Errorf("插件初始化失败: %w", err)
	}

	// 保存配置
	pm.Configs[pluginID] = config
	pm.Metrics.ActivePlugins++

	pm.Logger.Info("插件配置成功，插件ID: %s, 配置数量: %d", pluginID, len(config))

	return nil
}

// ExecutePlugin 执行插件
func (pm *PluginManager) ExecutePlugin(ctx context.Context, pluginID string, args map[string]interface{}) (*PluginResult, error) {
	pm.Mu.RLock()
	plugin, exists := pm.Plugins[pluginID]
	pm.Mu.RUnlock()

	if !exists {
		return nil, fmt.Errorf("插件 %s 不存在", pluginID)
	}

	startTime := time.Now()

	// 执行插件
	result, err := plugin.Execute(ctx, args)
	if result == nil {
		result = &PluginResult{}
	}

	executionTime := time.Since(startTime)

	// 设置结果
	result.ExecutionTime = executionTime
	result.Timestamp = time.Now()

	if err != nil {
		result.Success = false
		result.Error = err
		pm.Metrics.ErrorCount++
	} else {
		result.Success = true
		pm.Metrics.SuccessCount++
	}

	// 更新指标
	pm.Mu.Lock()
	pm.Metrics.TotalExecutions++
	pm.Metrics.LastExecution = time.Now()
	if pm.Metrics.TotalExecutions > 0 {
		totalTime := pm.Metrics.AverageTime * time.Duration(pm.Metrics.TotalExecutions-1)
		pm.Metrics.AverageTime = (totalTime + executionTime) / time.Duration(pm.Metrics.TotalExecutions)
	} else {
		pm.Metrics.AverageTime = executionTime
	}
	pm.Mu.Unlock()

	// 保存结果
	pm.Mu.Lock()
	pm.Results[pluginID] = result
	pm.Mu.Unlock()

	pm.Logger.Info("插件执行完成，插件ID: %s, 成功: %t, 执行时间: %v, 错误: %v", pluginID, result.Success, executionTime, err)

	return result, err
}

// GetPlugin 获取插件实例
func (pm *PluginManager) GetPlugin(pluginID string) (Plugin, error) {
	pm.Mu.RLock()
	defer pm.Mu.RUnlock()

	plugin, exists := pm.Plugins[pluginID]
	if !exists {
		return nil, fmt.Errorf("插件 %s 不存在", pluginID)
	}

	return plugin, nil
}

// GetPluginInfo 获取插件信息
func (pm *PluginManager) GetPluginInfo(pluginID string) (*PluginInfo, error) {
	pm.Mu.RLock()
	defer pm.Mu.RUnlock()

	info, exists := pm.PluginInfos[pluginID]
	if !exists {
		return nil, fmt.Errorf("插件 %s 不存在", pluginID)
	}

	return info, nil
}

// ListPlugins 列出所有插件
func (pm *PluginManager) ListPlugins() []*PluginInfo {
	pm.Mu.RLock()
	defer pm.Mu.RUnlock()

	infos := make([]*PluginInfo, 0, len(pm.PluginInfos))
	for _, info := range pm.PluginInfos {
		infos = append(infos, info)
	}

	return infos
}

// GetPluginResult 获取插件执行结果
func (pm *PluginManager) GetPluginResult(pluginID string) (*PluginResult, error) {
	pm.Mu.RLock()
	defer pm.Mu.RUnlock()

	result, exists := pm.Results[pluginID]
	if !exists {
		return nil, fmt.Errorf("插件 %s 的执行结果不存在", pluginID)
	}

	return result, nil
}

// GetMetrics 获取插件指标
func (pm *PluginManager) GetMetrics() *PluginMetrics {
	pm.Mu.RLock()
	defer pm.Mu.RUnlock()

	return pm.Metrics
}

// 辅助方法

// checkDependencies 检查插件依赖
func (pm *PluginManager) checkDependencies(info *PluginInfo) error {
	if info.Required == nil {
		return nil
	}
	
	for _, dep := range info.Required {
		if _, exists := pm.Plugins[dep]; !exists {
			return fmt.Errorf("缺少必需依赖: %s", dep)
		}
	}
	
	return nil
}
