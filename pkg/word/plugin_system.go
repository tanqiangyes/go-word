package word

import (
    "context"
    "encoding/json"
    "fmt"
    "os"
    "path/filepath"
    "plugin"
    "reflect"
    "strings"
    "sync"
    "time"

    "github.com/tanqiangyes/go-word/pkg/utils"
)

// PluginSystem 插件系统
type PluginSystem struct {
    Document    *Document
    Plugins     map[string]*Plugin
    Hooks       map[string][]PluginHook
    Registry    *PluginRegistry
    Config      *PluginConfig
    Logger      *utils.Logger
    Mu          sync.RWMutex
    Metrics     *PluginMetrics
    IsEnabled   bool
    LoadedPaths map[string]bool
}

// Plugin 插件
type Plugin struct {
    ID           string                 `json:"id"`
    Name         string                 `json:"name"`
    Version      string                 `json:"version"`
    Description  string                 `json:"description"`
    Author       string                 `json:"author"`
    License      string                 `json:"license"`
    Homepage     string                 `json:"homepage"`
    Type         PluginType             `json:"type"`
    Status       PluginStatus           `json:"status"`
    Path         string                 `json:"path"`
    EntryPoint   string                 `json:"entry_point"`
    Dependencies []string               `json:"dependencies"`
    Permissions  []string               `json:"permissions"`
    Hooks        []string               `json:"hooks"`
    Config       map[string]interface{} `json:"config"`
    Metadata     map[string]interface{} `json:"metadata"`
    Instance     interface{}            `json:"-"`
    LoadedAt     time.Time              `json:"loaded_at"`
    LastUsed     time.Time              `json:"last_used"`
}

// PluginType 插件类型
type PluginType string

const (
    PluginTypeProcessor  PluginType = "processor"  // 处理器插件
    PluginTypeFormatter  PluginType = "formatter"  // 格式化插件
    PluginTypeValidator  PluginType = "validator"  // 验证器插件
    PluginTypeExporter   PluginType = "exporter"   // 导出器插件
    PluginTypeImporter   PluginType = "importer"   // 导入器插件
    PluginTypeRenderer   PluginType = "renderer"   // 渲染器插件
    PluginTypeAnalyzer   PluginType = "analyzer"   // 分析器插件
    PluginTypeUtility    PluginType = "utility"    // 工具插件
    PluginTypeExtension  PluginType = "extension"  // 扩展插件
    PluginTypeMiddleware PluginType = "middleware" // 中间件插件
)

// PluginStatus 插件状态
type PluginStatus string

const (
    PluginStatusLoaded   PluginStatus = "loaded"   // 已加载
    PluginStatusActive   PluginStatus = "active"   // 活跃
    PluginStatusInactive PluginStatus = "inactive" // 非活跃
    PluginStatusError    PluginStatus = "error"    // 错误
    PluginStatusDisabled PluginStatus = "disabled" // 已禁用
    PluginStatusUnloaded PluginStatus = "unloaded" // 已卸载
)

// PluginHook 插件钩子
type PluginHook interface {
    Execute(ctx context.Context, args map[string]interface{}) (map[string]interface{}, error)
    GetPriority() int
    GetName() string
}

// PluginInterface 插件接口
type PluginInterface interface {
    Initialize(ctx context.Context, config map[string]interface{}) error
    Execute(ctx context.Context, args map[string]interface{}) (map[string]interface{}, error)
    Cleanup(ctx context.Context) error
    GetInfo() *PluginInfo
    GetHooks() []string
}

// PluginInfo 插件信息
type PluginInfo struct {
    ID          string   `json:"id"`
    Name        string   `json:"name"`
    Version     string   `json:"version"`
    Description string   `json:"description"`
    Author      string   `json:"author"`
    Hooks       []string `json:"hooks"`
}

// PluginRegistry 插件注册表
type PluginRegistry struct {
    Plugins      map[string]*Plugin
    Categories   map[PluginType][]*Plugin
    Dependencies map[string][]string
    Mu           sync.RWMutex
}

// PluginConfig 插件配置
type PluginConfig struct {
    Enabled        bool         `json:"enabled"`
    PluginDirs     []string     `json:"plugin_dirs"`
    MaxPlugins     int          `json:"max_plugins"`
    AutoLoad       bool         `json:"auto_load"`
    SandboxEnabled bool         `json:"sandbox_enabled"`
    TimeoutSeconds int          `json:"timeout_seconds"`
    AllowedTypes   []PluginType `json:"allowed_types"`
    Blacklist      []string     `json:"blacklist"`
    Whitelist      []string     `json:"whitelist"`
    SecurityLevel  string       `json:"security_level"`
    LogLevel       string       `json:"log_level"`
}

// PluginMetrics 插件指标
type PluginMetrics struct {
    TotalPlugins    int64         `json:"total_plugins"`
    ActivePlugins   int64         `json:"active_plugins"`
    LoadedPlugins   int64         `json:"loaded_plugins"`
    FailedPlugins   int64         `json:"failed_plugins"`
    HookExecutions  int64         `json:"hook_executions"`
    AverageLoadTime time.Duration `json:"average_load_time"`
    AverageExecTime time.Duration `json:"average_exec_time"`
    LastActivity    time.Time     `json:"last_activity"`
    ErrorCount      int64         `json:"error_count"`
    MemoryUsage     int64         `json:"memory_usage"`
}

// PluginEvent 插件事件
type PluginEvent struct {
    Type      string                 `json:"type"`
    PluginID  string                 `json:"plugin_id"`
    Hook      string                 `json:"hook"`
    Data      map[string]interface{} `json:"data"`
    Timestamp time.Time              `json:"timestamp"`
    Duration  time.Duration          `json:"duration"`
    Error     error                  `json:"error,omitempty"`
}

// PluginResult 插件执行结果
type PluginResult struct {
    Success  bool                   `json:"success"`
    PluginID string                 `json:"plugin_id"`
    Hook     string                 `json:"hook"`
    Data     map[string]interface{} `json:"data"`
    Duration time.Duration          `json:"duration"`
    Error    error                  `json:"error,omitempty"`
}

// NewPluginSystem 创建插件系统
func NewPluginSystem(document *Document, config *PluginConfig) *PluginSystem {
    if config == nil {
        config = getDefaultPluginConfig()
    }

    return &PluginSystem{
        Document:    document,
        Plugins:     make(map[string]*Plugin),
        Hooks:       make(map[string][]PluginHook),
        Registry:    NewPluginRegistry(),
        Config:      config,
        Logger:      utils.NewLogger(utils.LogLevelInfo, os.Stdout),
        Metrics:     &PluginMetrics{},
        IsEnabled:   config.Enabled,
        LoadedPaths: make(map[string]bool),
    }
}

// NewPluginRegistry 创建插件注册表
func NewPluginRegistry() *PluginRegistry {
    return &PluginRegistry{
        Plugins:      make(map[string]*Plugin),
        Categories:   make(map[PluginType][]*Plugin),
        Dependencies: make(map[string][]string),
    }
}

// LoadPlugin 加载插件
func (ps *PluginSystem) LoadPlugin(ctx context.Context, pluginPath string) (*Plugin, error) {
    ps.Mu.Lock()
    defer ps.Mu.Unlock()

    if !ps.IsEnabled {
        return nil, fmt.Errorf("plugin system is disabled")
    }

    // 检查是否已加载
    if ps.LoadedPaths[pluginPath] {
        return nil, fmt.Errorf("plugin already loaded: %s", pluginPath)
    }

    // 检查插件数量限制
    if len(ps.Plugins) >= ps.Config.MaxPlugins {
        return nil, fmt.Errorf("maximum number of plugins (%d) exceeded", ps.Config.MaxPlugins)
    }

    startTime := time.Now()

    // 加载插件文件
    p, err := plugin.Open(pluginPath)
    if err != nil {
        ps.Metrics.FailedPlugins++
        return nil, fmt.Errorf("failed to open plugin: %w", err)
    }

    // 查找插件符号
    symbol, err := p.Lookup("Plugin")
    if err != nil {
        ps.Metrics.FailedPlugins++
        return nil, fmt.Errorf("plugin symbol not found: %w", err)
    }

    // 类型断言
    pluginInterface, ok := symbol.(PluginInterface)
    if !ok {
        ps.Metrics.FailedPlugins++
        return nil, fmt.Errorf("invalid plugin interface")
    }

    // 获取插件信息
    info := pluginInterface.GetInfo()
    if info == nil {
        ps.Metrics.FailedPlugins++
        return nil, fmt.Errorf("plugin info is nil")
    }

    // 检查黑名单
    if ps.isBlacklisted(info.ID) {
        return nil, fmt.Errorf("plugin %s is blacklisted", info.ID)
    }

    // 检查白名单
    if len(ps.Config.Whitelist) > 0 && !ps.isWhitelisted(info.ID) {
        return nil, fmt.Errorf("plugin %s is not whitelisted", info.ID)
    }

    // 创建插件对象
    plugin := &Plugin{
        ID:          info.ID,
        Name:        info.Name,
        Version:     info.Version,
        Description: info.Description,
        Author:      info.Author,
        Type:        ps.detectPluginType(info),
        Status:      PluginStatusLoaded,
        Path:        pluginPath,
        Hooks:       info.Hooks,
        Config:      make(map[string]interface{}),
        Metadata:    make(map[string]interface{}),
        Instance:    pluginInterface,
        LoadedAt:    time.Now(),
        LastUsed:    time.Now(),
    }

    // 初始化插件
    err = pluginInterface.Initialize(ctx, plugin.Config)
    if err != nil {
        ps.Metrics.FailedPlugins++
        return nil, fmt.Errorf("failed to initialize plugin: %w", err)
    }

    // 注册插件
    ps.Plugins[plugin.ID] = plugin
    ps.LoadedPaths[pluginPath] = true
    ps.Registry.RegisterPlugin(plugin)

    // 注册钩子
    for _, hookName := range plugin.Hooks {
        ps.registerHook(hookName, pluginInterface)
    }

    // 更新指标
    loadTime := time.Since(startTime)
    ps.updateLoadMetrics(loadTime)

    plugin.Status = PluginStatusActive

    ps.Logger.Info("插件加载成功，插件ID: %s, 插件名称: %s, 版本: %s, 加载时间: %v, 钩子数: %d", plugin.ID, plugin.Name, plugin.Version, loadTime, len(plugin.Hooks))

    return plugin, nil
}

// UnloadPlugin 卸载插件
func (ps *PluginSystem) UnloadPlugin(ctx context.Context, pluginID string) error {
    ps.Mu.Lock()
    defer ps.Mu.Unlock()

    plugin, exists := ps.Plugins[pluginID]
    if !exists {
        return fmt.Errorf("plugin not found: %s", pluginID)
    }

    // 清理插件
    if pluginInterface, ok := plugin.Instance.(PluginInterface); ok {
        err := pluginInterface.Cleanup(ctx)
        if err != nil {
            ps.Logger.Warning("插件清理失败，插件ID: %s, 错误: %s", pluginID, err.Error())
        }
    }

    // 注销钩子
    for _, hookName := range plugin.Hooks {
        ps.unregisterHook(hookName, plugin.Instance)
    }

    // 从注册表移除
    ps.Registry.UnregisterPlugin(pluginID)

    // 移除插件
    delete(ps.Plugins, pluginID)
    delete(ps.LoadedPaths, plugin.Path)

    // 更新指标
    ps.Metrics.LoadedPlugins--
    if plugin.Status == PluginStatusActive {
        ps.Metrics.ActivePlugins--
    }

    ps.Logger.Info("插件卸载成功，插件ID: %s, 插件名称: %s", pluginID, plugin.Name)

    return nil
}

// ExecuteHook 执行钩子
func (ps *PluginSystem) ExecuteHook(ctx context.Context, hookName string, args map[string]interface{}) ([]*PluginResult, error) {
    ps.Mu.RLock()
    hooks, exists := ps.Hooks[hookName]
    ps.Mu.RUnlock()

    if !exists || len(hooks) == 0 {
        return nil, fmt.Errorf("no hooks registered for: %s", hookName)
    }

    results := make([]*PluginResult, 0, len(hooks))

    // 按优先级排序执行
    for _, hook := range hooks {
        startTime := time.Now()

        result := &PluginResult{
            Hook: hookName,
        }

        // 执行钩子
        data, err := hook.Execute(ctx, args)
        duration := time.Since(startTime)

        result.Duration = duration
        result.Data = data

        if err != nil {
            result.Error = err
            ps.Metrics.ErrorCount++
            ps.Logger.Error("钩子执行失败，钩子: %s, 错误: %s, 执行时间: %v", hookName, err.Error(), duration)
        } else {
            result.Success = true
        }

        results = append(results, result)

        // 更新指标
        ps.Mu.Lock()
        ps.Metrics.HookExecutions++
        ps.Metrics.LastActivity = time.Now()
        ps.updateExecMetrics(duration)
        ps.Mu.Unlock()
    }

    return results, nil
}

// GetPlugin 获取插件
func (ps *PluginSystem) GetPlugin(pluginID string) (*Plugin, error) {
    ps.Mu.RLock()
    defer ps.Mu.RUnlock()

    plugin, exists := ps.Plugins[pluginID]
    if !exists {
        return nil, fmt.Errorf("plugin not found: %s", pluginID)
    }

    return plugin, nil
}

// ListPlugins 列出所有插件
func (ps *PluginSystem) ListPlugins() []*Plugin {
    ps.Mu.RLock()
    defer ps.Mu.RUnlock()

    plugins := make([]*Plugin, 0, len(ps.Plugins))
    for _, plugin := range ps.Plugins {
        plugins = append(plugins, plugin)
    }

    return plugins
}

// ListPluginsByType 按类型列出插件
func (ps *PluginSystem) ListPluginsByType(pluginType PluginType) []*Plugin {
    ps.Mu.RLock()
    defer ps.Mu.RUnlock()

    return ps.Registry.GetPluginsByType(pluginType)
}

// EnablePlugin 启用插件
func (ps *PluginSystem) EnablePlugin(pluginID string) error {
    ps.Mu.Lock()
    defer ps.Mu.Unlock()

    plugin, exists := ps.Plugins[pluginID]
    if !exists {
        return fmt.Errorf("plugin not found: %s", pluginID)
    }

    if plugin.Status == PluginStatusActive {
        return nil // 已经是活跃状态
    }

    oldStatus := plugin.Status
    plugin.Status = PluginStatusActive

    if oldStatus != PluginStatusActive {
        ps.Metrics.ActivePlugins++
    }

    ps.Logger.Info("插件已启用，插件ID: %s, 旧状态: %s", pluginID, oldStatus)

    return nil
}

// DisablePlugin 禁用插件
func (ps *PluginSystem) DisablePlugin(pluginID string) error {
    ps.Mu.Lock()
    defer ps.Mu.Unlock()

    plugin, exists := ps.Plugins[pluginID]
    if !exists {
        return fmt.Errorf("plugin not found: %s", pluginID)
    }

    if plugin.Status == PluginStatusDisabled {
        return nil // 已经是禁用状态
    }

    oldStatus := plugin.Status
    plugin.Status = PluginStatusDisabled

    if oldStatus == PluginStatusActive {
        ps.Metrics.ActivePlugins--
    }

    ps.Logger.Info("插件已禁用，插件ID: %s, 旧状态: %s", pluginID, oldStatus)

    return nil
}

// LoadPluginsFromDir 从目录加载插件
func (ps *PluginSystem) LoadPluginsFromDir(ctx context.Context, dir string) error {
    if !ps.Config.AutoLoad {
        return fmt.Errorf("auto load is disabled")
    }

    // 遍历目录
    err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
        if err != nil {
            return err
        }

        // 只处理.so文件（Linux）或.dll文件（Windows）
        if !info.IsDir() && (strings.HasSuffix(path, ".so") || strings.HasSuffix(path, ".dll")) {
            _, err := ps.LoadPlugin(ctx, path)
            if err != nil {
                ps.Logger.Warning("插件加载失败，路径: %s, 错误: %s", path, err.Error())
            }
        }

        return nil
    })

    if err != nil {
        return fmt.Errorf("failed to load plugins from directory: %w", err)
    }

    ps.Logger.Info("目录插件加载完成，目录: %s, 已加载: %d", dir, len(ps.Plugins))

    return nil
}

// GetMetrics 获取指标
func (ps *PluginSystem) GetMetrics() *PluginMetrics {
    ps.Mu.RLock()
    defer ps.Mu.RUnlock()

    return ps.Metrics
}

// ExportPluginConfig 导出插件配置
func (ps *PluginSystem) ExportPluginConfig() ([]byte, error) {
    ps.Mu.RLock()
    defer ps.Mu.RUnlock()

    type PluginExport struct {
        Plugins []*Plugin      `json:"plugins"`
        Config  *PluginConfig  `json:"config"`
        Metrics *PluginMetrics `json:"metrics"`
    }

    export := &PluginExport{
        Plugins: ps.ListPlugins(),
        Config:  ps.Config,
        Metrics: ps.Metrics,
    }

    data, err := json.MarshalIndent(export, "", "  ")
    if err != nil {
        return nil, fmt.Errorf("failed to export plugin config: %w", err)
    }

    return data, nil
}

// 插件注册表方法

// RegisterPlugin 注册插件
func (pr *PluginRegistry) RegisterPlugin(plugin *Plugin) {
    pr.Mu.Lock()
    defer pr.Mu.Unlock()

    pr.Plugins[plugin.ID] = plugin

    // 按类型分类
    if pr.Categories[plugin.Type] == nil {
        pr.Categories[plugin.Type] = make([]*Plugin, 0)
    }
    pr.Categories[plugin.Type] = append(pr.Categories[plugin.Type], plugin)

    // 记录依赖关系
    pr.Dependencies[plugin.ID] = plugin.Dependencies
}

// UnregisterPlugin 注销插件
func (pr *PluginRegistry) UnregisterPlugin(pluginID string) {
    pr.Mu.Lock()
    defer pr.Mu.Unlock()

    plugin, exists := pr.Plugins[pluginID]
    if !exists {
        return
    }

    // 从分类中移除
    category := pr.Categories[plugin.Type]
    for i, p := range category {
        if p.ID == pluginID {
            pr.Categories[plugin.Type] = append(category[:i], category[i+1:]...)
            break
        }
    }

    // 移除插件和依赖
    delete(pr.Plugins, pluginID)
    delete(pr.Dependencies, pluginID)
}

// GetPluginsByType 按类型获取插件
func (pr *PluginRegistry) GetPluginsByType(pluginType PluginType) []*Plugin {
    pr.Mu.RLock()
    defer pr.Mu.RUnlock()

    return pr.Categories[pluginType]
}

// 辅助方法

// registerHook 注册钩子
func (ps *PluginSystem) registerHook(hookName string, pluginInterface interface{}) {
    if hook, ok := pluginInterface.(PluginHook); ok {
        if ps.Hooks[hookName] == nil {
            ps.Hooks[hookName] = make([]PluginHook, 0)
        }
        ps.Hooks[hookName] = append(ps.Hooks[hookName], hook)
    }
}

// unregisterHook 注销钩子
func (ps *PluginSystem) unregisterHook(hookName string, pluginInterface interface{}) {
    hooks := ps.Hooks[hookName]
    for i, hook := range hooks {
        if reflect.DeepEqual(hook, pluginInterface) {
            ps.Hooks[hookName] = append(hooks[:i], hooks[i+1:]...)
            break
        }
    }
}

// detectPluginType 检测插件类型
func (ps *PluginSystem) detectPluginType(info *PluginInfo) PluginType {
    // 根据插件名称或钩子推断类型
    name := strings.ToLower(info.Name)

    if strings.Contains(name, "processor") {
        return PluginTypeProcessor
    }
    if strings.Contains(name, "formatter") {
        return PluginTypeFormatter
    }
    if strings.Contains(name, "validator") {
        return PluginTypeValidator
    }
    if strings.Contains(name, "exporter") {
        return PluginTypeExporter
    }
    if strings.Contains(name, "importer") {
        return PluginTypeImporter
    }

    return PluginTypeExtension // 默认类型
}

// isBlacklisted 检查是否在黑名单中
func (ps *PluginSystem) isBlacklisted(pluginID string) bool {
    for _, blacklisted := range ps.Config.Blacklist {
        if pluginID == blacklisted {
            return true
        }
    }
    return false
}

// isWhitelisted 检查是否在白名单中
func (ps *PluginSystem) isWhitelisted(pluginID string) bool {
    for _, whitelisted := range ps.Config.Whitelist {
        if pluginID == whitelisted {
            return true
        }
    }
    return false
}

// updateLoadMetrics 更新加载指标
func (ps *PluginSystem) updateLoadMetrics(loadTime time.Duration) {
    ps.Metrics.TotalPlugins++
    ps.Metrics.LoadedPlugins++
    ps.Metrics.ActivePlugins++

    if ps.Metrics.TotalPlugins > 0 {
        ps.Metrics.AverageLoadTime = time.Duration(
            (int64(ps.Metrics.AverageLoadTime)*int64(ps.Metrics.TotalPlugins-1) + int64(loadTime)) / int64(ps.Metrics.TotalPlugins),
        )
    } else {
        ps.Metrics.AverageLoadTime = loadTime
    }
}

// updateExecMetrics 更新执行指标
func (ps *PluginSystem) updateExecMetrics(execTime time.Duration) {
    if ps.Metrics.HookExecutions > 0 {
        ps.Metrics.AverageExecTime = time.Duration(
            (int64(ps.Metrics.AverageExecTime)*int64(ps.Metrics.HookExecutions-1) + int64(execTime)) / int64(ps.Metrics.HookExecutions),
        )
    } else {
        ps.Metrics.AverageExecTime = execTime
    }
}

// getDefaultPluginConfig 获取默认配置
func getDefaultPluginConfig() *PluginConfig {
    return &PluginConfig{
        Enabled:        true,
        PluginDirs:     []string{"./plugins", "./extensions"},
        MaxPlugins:     100,
        AutoLoad:       true,
        SandboxEnabled: true,
        TimeoutSeconds: 30,
        AllowedTypes: []PluginType{
            PluginTypeProcessor,
            PluginTypeFormatter,
            PluginTypeValidator,
            PluginTypeExporter,
            PluginTypeImporter,
            PluginTypeRenderer,
            PluginTypeAnalyzer,
            PluginTypeUtility,
            PluginTypeExtension,
        },
        Blacklist:     []string{},
        Whitelist:     []string{},
        SecurityLevel: "medium",
        LogLevel:      "info",
    }
}
