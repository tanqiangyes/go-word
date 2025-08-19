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
	ID          string            `json:"id"`
	Name        string            `json:"name"`
	Version     string            `json:"version"`
	Description string            `json:"description"`
	Author      string            `json:"author"`
	License     string            `json:"license"`
	Category    string            `json:"category"`
	Tags        []string          `json:"tags"`
	Required    []string          `json:"required"`
	Optional    []string          `json:"optional"`
	CreatedAt   time.Time         `json:"created_at"`
	UpdatedAt   time.Time         `json:"updated_at"`
	Metadata    map[string]interface{} `json:"metadata"`
}

// PluginResult 插件执行结果
type PluginResult struct {
	Success     bool                   `json:"success"`
	Data        map[string]interface{} `json:"data"`
	Message     string                 `json:"message"`
	Error       error                  `json:"error,omitempty"`
	ExecutionTime time.Duration       `json:"execution_time"`
	Timestamp   time.Time             `json:"timestamp"`
}

// PluginManager 插件管理器
type PluginManager struct {
	plugins       map[string]Plugin
	pluginInfos   map[string]*PluginInfo
	configs       map[string]map[string]interface{}
	results       map[string]*PluginResult
	mu            sync.RWMutex
	logger        *utils.Logger
	metrics       *PluginMetrics
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
		plugins:     make(map[string]Plugin),
		pluginInfos: make(map[string]*PluginInfo),
		configs:     make(map[string]map[string]interface{}),
		results:     make(map[string]*PluginResult),
		logger:      utils.NewLogger(utils.LogLevelInfo, nil),
		metrics:     &PluginMetrics{},
	}
}

// RegisterPlugin 注册插件
func (pm *PluginManager) RegisterPlugin(plugin Plugin) error {
	pm.mu.Lock()
	defer pm.mu.Unlock()
	
	info := plugin.GetInfo()
	if info == nil {
		return fmt.Errorf("插件信息不能为空")
	}
	
	if info.ID == "" {
		return fmt.Errorf("插件ID不能为空")
	}
	
	if _, exists := pm.plugins[info.ID]; exists {
		return fmt.Errorf("插件 %s 已存在", info.ID)
	}
	
	// 检查依赖
	if err := pm.checkDependencies(info); err != nil {
		return fmt.Errorf("依赖检查失败: %w", err)
	}
	
	// 注册插件
	pm.plugins[info.ID] = plugin
	pm.pluginInfos[info.ID] = info
	pm.metrics.TotalPlugins++
	
	pm.logger.Info("插件注册成功", map[string]interface{}{
		"plugin_id":   info.ID,
		"name":        info.Name,
		"version":     info.Version,
		"category":    info.Category,
	})
	
	return nil
}

// UnregisterPlugin 注销插件
func (pm *PluginManager) UnregisterPlugin(pluginID string) error {
	pm.mu.Lock()
	defer pm.mu.Unlock()
	
	plugin, exists := pm.plugins[pluginID]
	if !exists {
		return fmt.Errorf("插件 %s 不存在", pluginID)
	}
	
	// 清理插件资源
	if err := plugin.Cleanup(); err != nil {
		pm.logger.Warning("插件清理失败", map[string]interface{}{
			"plugin_id": pluginID,
			"error":     err.Error(),
		})
	}
	
	// 注销插件
	delete(pm.plugins, pluginID)
	delete(pm.pluginInfos, pluginID)
	delete(pm.configs, pluginID)
	delete(pm.results, pluginID)
	
	pm.metrics.TotalPlugins--
	
	pm.logger.Info("插件注销成功", map[string]interface{}{
		"plugin_id": pluginID,
	})
	
	return nil
}

// ConfigurePlugin 配置插件
func (pm *PluginManager) ConfigurePlugin(pluginID string, config map[string]interface{}) error {
	pm.mu.Lock()
	defer pm.mu.Unlock()
	
	plugin, exists := pm.plugins[pluginID]
	if !exists {
		return fmt.Errorf("插件 %s 不存在", pluginID)
	}
	
	// 初始化插件
	if err := plugin.Initialize(config); err != nil {
		return fmt.Errorf("插件初始化失败: %w", err)
	}
	
	// 保存配置
	pm.configs[pluginID] = config
	pm.metrics.ActivePlugins++
	
	pm.logger.Info("插件配置成功", map[string]interface{}{
		"plugin_id": pluginID,
		"config":    config,
	})
	
	return nil
}

// ExecutePlugin 执行插件
func (pm *PluginManager) ExecutePlugin(ctx context.Context, pluginID string, args map[string]interface{}) (*PluginResult, error) {
	pm.mu.RLock()
	plugin, exists := pm.plugins[pluginID]
	pm.mu.RUnlock()
	
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
		pm.metrics.ErrorCount++
	} else {
		result.Success = true
		pm.metrics.SuccessCount++
	}
	
	// 更新指标
	pm.mu.Lock()
	pm.metrics.TotalExecutions++
	pm.metrics.LastExecution = time.Now()
	if pm.metrics.TotalExecutions > 0 {
		totalTime := pm.metrics.AverageTime * time.Duration(pm.metrics.TotalExecutions-1)
		pm.metrics.AverageTime = (totalTime + executionTime) / time.Duration(pm.metrics.TotalExecutions)
	} else {
		pm.metrics.AverageTime = executionTime
	}
	pm.mu.Unlock()
	
	// 保存结果
	pm.mu.Lock()
	pm.results[pluginID] = result
	pm.mu.Unlock()
	
	pm.logger.Info("插件执行完成", map[string]interface{}{
		"plugin_id":     pluginID,
		"success":        result.Success,
		"execution_time": executionTime,
		"error":          err,
	})
	
	return result, err
}

// GetPlugin 获取插件
func (pm *PluginManager) GetPlugin(pluginID string) (Plugin, error) {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	
	plugin, exists := pm.plugins[pluginID]
	if !exists {
		return nil, fmt.Errorf("插件 %s 不存在", pluginID)
	}
	
	return plugin, nil
}

// GetPluginInfo 获取插件信息
func (pm *PluginManager) GetPluginInfo(pluginID string) (*PluginInfo, error) {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	
	info, exists := pm.pluginInfos[pluginID]
	if !exists {
		return nil, fmt.Errorf("插件 %s 不存在", pluginID)
	}
	
	return info, nil
}

// ListPlugins 列出所有插件
func (pm *PluginManager) ListPlugins() []*PluginInfo {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	
	plugins := make([]*PluginInfo, 0, len(pm.pluginInfos))
	for _, info := range pm.pluginInfos {
		plugins = append(plugins, info)
	}
	
	return plugins
}

// GetPluginResult 获取插件执行结果
func (pm *PluginManager) GetPluginResult(pluginID string) (*PluginResult, error) {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	
	result, exists := pm.results[pluginID]
	if !exists {
		return nil, fmt.Errorf("插件 %s 的执行结果不存在", pluginID)
	}
	
	return result, nil
}

// GetMetrics 获取插件指标
func (pm *PluginManager) GetMetrics() *PluginMetrics {
	pm.mu.RLock()
	defer pm.mu.RUnlock()
	
	return pm.metrics
}

// 辅助方法

// checkDependencies 检查插件依赖
func (pm *PluginManager) checkDependencies(info *PluginInfo) error {
	if len(info.Required) == 0 {
		return nil
	}
	
	for _, dep := range info.Required {
		if _, exists := pm.plugins[dep]; !exists {
			return fmt.Errorf("缺少必需依赖: %s", dep)
		}
	}
	
	return nil
}
