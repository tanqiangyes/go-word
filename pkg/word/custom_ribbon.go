package word

import (
    "context"
    "encoding/xml"
    "fmt"
    "os"
    "sync"
    "time"

    "github.com/tanqiangyes/go-word/pkg/utils"
)

// CustomRibbon 自定义功能区
type CustomRibbon struct {
    document  *Document
    tabs      map[string]*RibbonTab
    groups    map[string]*RibbonGroup
    controls  map[string]*RibbonControl
    callbacks map[string]RibbonCallback
    config    *RibbonConfig
    logger    *utils.Logger
    mu        sync.RWMutex
    metrics   *RibbonMetrics
    isEnabled bool
}

// RibbonTab 功能区选项卡
type RibbonTab struct {
    ID          string                 `json:"id" xml:"id,attr"`
    Label       string                 `json:"label" xml:"label,attr"`
    Description string                 `json:"description" xml:"description,attr"`
    Visible     bool                   `json:"visible" xml:"visible,attr"`
    Enabled     bool                   `json:"enabled" xml:"enabled,attr"`
    Position    int                    `json:"position" xml:"position,attr"`
    Groups      []*RibbonGroup         `json:"groups" xml:"group"`
    Properties  map[string]interface{} `json:"properties"`
    CreatedAt   time.Time              `json:"created_at"`
    UpdatedAt   time.Time              `json:"updated_at"`
}

// RibbonGroup 功能区组
type RibbonGroup struct {
    ID          string                 `json:"id" xml:"id,attr"`
    Label       string                 `json:"label" xml:"label,attr"`
    Description string                 `json:"description" xml:"description,attr"`
    Visible     bool                   `json:"visible" xml:"visible,attr"`
    Enabled     bool                   `json:"enabled" xml:"enabled,attr"`
    Position    int                    `json:"position" xml:"position,attr"`
    TabID       string                 `json:"tab_id" xml:"tab_id,attr"`
    Controls    []*RibbonControl       `json:"controls" xml:"control"`
    Properties  map[string]interface{} `json:"properties"`
    CreatedAt   time.Time              `json:"created_at"`
    UpdatedAt   time.Time              `json:"updated_at"`
}

// RibbonControl 功能区控件
type RibbonControl struct {
    ID          string                 `json:"id" xml:"id,attr"`
    Type        ControlType            `json:"type" xml:"type,attr"`
    Label       string                 `json:"label" xml:"label,attr"`
    Description string                 `json:"description" xml:"description,attr"`
    Tooltip     string                 `json:"tooltip" xml:"tooltip,attr"`
    Icon        string                 `json:"icon" xml:"icon,attr"`
    Visible     bool                   `json:"visible" xml:"visible,attr"`
    Enabled     bool                   `json:"enabled" xml:"enabled,attr"`
    Position    int                    `json:"position" xml:"position,attr"`
    GroupID     string                 `json:"group_id" xml:"group_id,attr"`
    Action      string                 `json:"action" xml:"action,attr"`
    Shortcut    string                 `json:"shortcut" xml:"shortcut,attr"`
    Size        ControlSize            `json:"size" xml:"size,attr"`
    Style       *ControlStyle          `json:"style" xml:"style"`
    Options     *ControlOptions        `json:"options" xml:"options"`
    Properties  map[string]interface{} `json:"properties"`
    CreatedAt   time.Time              `json:"created_at"`
    UpdatedAt   time.Time              `json:"updated_at"`
}

// ControlType 控件类型
type ControlType string

const (
    ControlTypeButton       ControlType = "button"      // 按钮
    ControlTypeToggleButton ControlType = "toggle"      // 切换按钮
    ControlTypeDropDown     ControlType = "dropdown"    // 下拉菜单
    ControlTypeComboBox     ControlType = "combobox"    // 组合框
    ControlTypeTextBox      ControlType = "textbox"     // 文本框
    ControlTypeSpinner      ControlType = "spinner"     // 数字调节器
    ControlTypeCheckBox     ControlType = "checkbox"    // 复选框
    ControlTypeRadioButton  ControlType = "radio"       // 单选按钮
    ControlTypeSeparator    ControlType = "separator"   // 分隔符
    ControlTypeLabel        ControlType = "label"       // 标签
    ControlTypeGallery      ControlType = "gallery"     // 图库
    ControlTypeMenu         ControlType = "menu"        // 菜单
    ControlTypeSplitButton  ControlType = "splitbutton" // 分割按钮
)

// ControlSize 控件大小
type ControlSize string

const (
    ControlSizeSmall  ControlSize = "small"  // 小
    ControlSizeMedium ControlSize = "medium" // 中
    ControlSizeLarge  ControlSize = "large"  // 大
)

// ControlStyle 控件样式
type ControlStyle struct {
    BackgroundColor string `json:"background_color" xml:"background_color,attr"`
    ForegroundColor string `json:"foreground_color" xml:"foreground_color,attr"`
    BorderColor     string `json:"border_color" xml:"border_color,attr"`
    BorderWidth     int    `json:"border_width" xml:"border_width,attr"`
    FontFamily      string `json:"font_family" xml:"font_family,attr"`
    FontSize        int    `json:"font_size" xml:"font_size,attr"`
    FontWeight      string `json:"font_weight" xml:"font_weight,attr"`
    Padding         int    `json:"padding" xml:"padding,attr"`
    Margin          int    `json:"margin" xml:"margin,attr"`
}

// ControlOptions 控件选项
type ControlOptions struct {
    Items        []string               `json:"items" xml:"item"`
    DefaultValue string                 `json:"default_value" xml:"default_value,attr"`
    MinValue     *float64               `json:"min_value" xml:"min_value,attr"`
    MaxValue     *float64               `json:"max_value" xml:"max_value,attr"`
    Step         *float64               `json:"step" xml:"step,attr"`
    MultiSelect  bool                   `json:"multi_select" xml:"multi_select,attr"`
    Editable     bool                   `json:"editable" xml:"editable,attr"`
    Properties   map[string]interface{} `json:"properties"`
}

// RibbonCallback 功能区回调函数
type RibbonCallback func(ctx context.Context, control *RibbonControl, args map[string]interface{}) error

// RibbonConfig 功能区配置
type RibbonConfig struct {
    Enabled             bool          `json:"enabled"`
    MaxTabs             int           `json:"max_tabs"`
    MaxGroupsPerTab     int           `json:"max_groups_per_tab"`
    MaxControlsPerGroup int           `json:"max_controls_per_group"`
    AllowedControlTypes []ControlType `json:"allowed_control_types"`
    Theme               string        `json:"theme"`
    AnimationEnabled    bool          `json:"animation_enabled"`
    TooltipEnabled      bool          `json:"tooltip_enabled"`
    ShortcutEnabled     bool          `json:"shortcut_enabled"`
    AutoSave            bool          `json:"auto_save"`
    AutoSaveInterval    int           `json:"auto_save_interval"`
}

// RibbonMetrics 功能区指标
type RibbonMetrics struct {
    TotalTabs       int64         `json:"total_tabs"`
    TotalGroups     int64         `json:"total_groups"`
    TotalControls   int64         `json:"total_controls"`
    ActiveControls  int64         `json:"active_controls"`
    CallbackCount   int64         `json:"callback_count"`
    LastInteraction time.Time     `json:"last_interaction"`
    AverageResponse time.Duration `json:"average_response"`
    ErrorCount      int64         `json:"error_count"`
}

// RibbonEvent 功能区事件
type RibbonEvent struct {
    Type      string                 `json:"type"`
    ControlID string                 `json:"control_id"`
    Data      map[string]interface{} `json:"data"`
    Timestamp time.Time              `json:"timestamp"`
}

// NewCustomRibbon 创建自定义功能区
func NewCustomRibbon(document *Document, config *RibbonConfig) *CustomRibbon {
    if config == nil {
        config = getDefaultRibbonConfig()
    }

    return &CustomRibbon{
        document:  document,
        config:    config,
        tabs:      make(map[string]*RibbonTab),
        groups:    make(map[string]*RibbonGroup),
        controls:  make(map[string]*RibbonControl),
        logger:    utils.NewLogger(utils.LogLevelInfo, os.Stdout),
        metrics:   &RibbonMetrics{},
        isEnabled: config.Enabled,
    }
}

// AddTab 添加选项卡
func (cr *CustomRibbon) AddTab(tab *RibbonTab) error {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    if !cr.isEnabled {
        return fmt.Errorf("custom ribbon is disabled")
    }

    // 检查选项卡数量限制
    if len(cr.tabs) >= cr.config.MaxTabs {
        return fmt.Errorf("maximum number of tabs (%d) exceeded", cr.config.MaxTabs)
    }

    // 检查ID是否已存在
    if _, exists := cr.tabs[tab.ID]; exists {
        return fmt.Errorf("tab with ID %s already exists", tab.ID)
    }

    // 设置默认值
    if tab.Position == 0 {
        tab.Position = len(cr.tabs) + 1
    }
    tab.CreatedAt = time.Now()
    tab.UpdatedAt = time.Now()

    if tab.Properties == nil {
        tab.Properties = make(map[string]interface{})
    }

    // 添加选项卡
    cr.tabs[tab.ID] = tab
    cr.metrics.TotalTabs++

    cr.logger.Info("选项卡添加成功，选项卡ID: %s, 标签: %s, 位置: %d", tab.ID, tab.Label, tab.Position)

    return nil
}

// AddGroup 添加组
func (cr *CustomRibbon) AddGroup(group *RibbonGroup) error {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    if !cr.isEnabled {
        return fmt.Errorf("custom ribbon is disabled")
    }

    // 检查选项卡是否存在
    tab, exists := cr.tabs[group.TabID]
    if !exists {
        return fmt.Errorf("tab with ID %s not found", group.TabID)
    }

    // 检查组数量限制
    if len(tab.Groups) >= cr.config.MaxGroupsPerTab {
        return fmt.Errorf("maximum number of groups per tab (%d) exceeded", cr.config.MaxGroupsPerTab)
    }

    // 检查ID是否已存在
    if _, exists := cr.groups[group.ID]; exists {
        return fmt.Errorf("group with ID %s already exists", group.ID)
    }

    // 设置默认值
    if group.Position == 0 {
        group.Position = len(tab.Groups) + 1
    }
    group.CreatedAt = time.Now()
    group.UpdatedAt = time.Now()

    if group.Properties == nil {
        group.Properties = make(map[string]interface{})
    }

    // 添加组
    cr.groups[group.ID] = group
    tab.Groups = append(tab.Groups, group)
    cr.metrics.TotalGroups++

    cr.logger.Info("组添加成功，组ID: %s, 标签: %s, 选项卡ID: %s, 位置: %d", group.ID, group.Label, group.TabID, group.Position)

    return nil
}

// AddControl 添加控件
func (cr *CustomRibbon) AddControl(control *RibbonControl) error {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    if !cr.isEnabled {
        return fmt.Errorf("custom ribbon is disabled")
    }

    // 检查组是否存在
    group, exists := cr.groups[control.GroupID]
    if !exists {
        return fmt.Errorf("group with ID %s not found", control.GroupID)
    }

    // 检查控件数量限制
    if len(group.Controls) >= cr.config.MaxControlsPerGroup {
        return fmt.Errorf("maximum number of controls per group (%d) exceeded", cr.config.MaxControlsPerGroup)
    }

    // 检查控件类型是否允许
    if !cr.isControlTypeAllowed(control.Type) {
        return fmt.Errorf("control type %s is not allowed", control.Type)
    }

    // 检查ID是否已存在
    if _, exists := cr.controls[control.ID]; exists {
        return fmt.Errorf("control with ID %s already exists", control.ID)
    }

    // 设置默认值
    if control.Position == 0 {
        control.Position = len(group.Controls) + 1
    }
    if control.Size == "" {
        control.Size = ControlSizeMedium
    }
    control.CreatedAt = time.Now()
    control.UpdatedAt = time.Now()

    if control.Properties == nil {
        control.Properties = make(map[string]interface{})
    }

    // 设置默认样式
    if control.Style == nil {
        control.Style = cr.getDefaultControlStyle()
    }

    // 添加控件
    cr.controls[control.ID] = control
    group.Controls = append(group.Controls, control)
    cr.metrics.TotalControls++

    if control.Enabled {
        cr.metrics.ActiveControls++
    }

    cr.logger.Info("控件添加成功，控件ID: %s, 类型: %s, 标签: %s, 组ID: %s, 位置: %d", control.ID, control.Type, control.Label, control.GroupID, control.Position)

    return nil
}

// RegisterCallback 注册回调函数
func (cr *CustomRibbon) RegisterCallback(controlID string, callback RibbonCallback) error {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    // 检查控件是否存在
    if _, exists := cr.controls[controlID]; !exists {
        return fmt.Errorf("control with ID %s not found", controlID)
    }

    cr.callbacks[controlID] = callback

    cr.logger.Info("回调函数注册成功，控件ID: %s", controlID)

    return nil
}

// TriggerControl 触发控件
func (cr *CustomRibbon) TriggerControl(ctx context.Context, controlID string, args map[string]interface{}) error {
    cr.mu.RLock()
    control, controlExists := cr.controls[controlID]
    callback, callbackExists := cr.callbacks[controlID]
    cr.mu.RUnlock()

    if !controlExists {
        return fmt.Errorf("control with ID %s not found", controlID)
    }

    if !control.Enabled {
        return fmt.Errorf("control %s is disabled", controlID)
    }

    if !callbackExists {
        return fmt.Errorf("no callback registered for control %s", controlID)
    }

    startTime := time.Now()

    // 执行回调
    err := callback(ctx, control, args)

    processTime := time.Since(startTime)

    // 更新指标
    cr.mu.Lock()
    cr.metrics.CallbackCount++
    cr.metrics.LastInteraction = time.Now()
    if cr.metrics.CallbackCount > 0 {
        cr.metrics.AverageResponse = time.Duration(int64(cr.metrics.AverageResponse)*int64(cr.metrics.CallbackCount-1)+int64(processTime)) / time.Duration(cr.metrics.CallbackCount)
    } else {
        cr.metrics.AverageResponse = processTime
    }

    if err != nil {
        cr.metrics.ErrorCount++
    }
    cr.mu.Unlock()

    if err != nil {
        cr.logger.Error("控件触发失败，控件ID: %s, 错误: %s, 处理时间: %v", controlID, err.Error(), processTime)
        return err
    }

    cr.logger.Info("控件触发成功，控件ID: %s, 处理时间: %v", controlID, processTime)

    return nil
}

// UpdateControl 更新控件
func (cr *CustomRibbon) UpdateControl(controlID string, updates map[string]interface{}) error {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    control, exists := cr.controls[controlID]
    if !exists {
        return fmt.Errorf("control with ID %s not found", controlID)
    }

    // 更新控件属性
    for key, value := range updates {
        switch key {
        case "label":
            if label, ok := value.(string); ok {
                control.Label = label
            }
        case "description":
            if desc, ok := value.(string); ok {
                control.Description = desc
            }
        case "tooltip":
            if tooltip, ok := value.(string); ok {
                control.Tooltip = tooltip
            }
        case "visible":
            if visible, ok := value.(bool); ok {
                control.Visible = visible
            }
        case "enabled":
            if enabled, ok := value.(bool); ok {
                oldEnabled := control.Enabled
                control.Enabled = enabled
                // 更新活跃控件计数
                if oldEnabled && !enabled {
                    cr.metrics.ActiveControls--
                } else if !oldEnabled && enabled {
                    cr.metrics.ActiveControls++
                }
            }
        case "icon":
            if icon, ok := value.(string); ok {
                control.Icon = icon
            }
        default:
            // 存储到属性中
            control.Properties[key] = value
        }
    }

    control.UpdatedAt = time.Now()

    cr.logger.Info("控件更新成功，控件ID: %s, 更新数量: %d", controlID, len(updates))

    return nil
}

// RemoveControl 移除控件
func (cr *CustomRibbon) RemoveControl(controlID string) error {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    control, exists := cr.controls[controlID]
    if !exists {
        return fmt.Errorf("control with ID %s not found", controlID)
    }

    // 从组中移除控件
    group := cr.groups[control.GroupID]
    for i, c := range group.Controls {
        if c.ID == controlID {
            group.Controls = append(group.Controls[:i], group.Controls[i+1:]...)
            break
        }
    }

    // 移除控件和回调
    delete(cr.controls, controlID)
    delete(cr.callbacks, controlID)

    // 更新指标
    cr.metrics.TotalControls--
    if control.Enabled {
        cr.metrics.ActiveControls--
    }

    cr.logger.Info("控件移除成功，控件ID: %s, 标签: %s", controlID, control.Label)

    return nil
}

// GetTab 获取选项卡
func (cr *CustomRibbon) GetTab(tabID string) (*RibbonTab, error) {
    cr.mu.RLock()
    defer cr.mu.RUnlock()

    tab, exists := cr.tabs[tabID]
    if !exists {
        return nil, fmt.Errorf("tab with ID %s not found", tabID)
    }

    return tab, nil
}

// GetGroup 获取组
func (cr *CustomRibbon) GetGroup(groupID string) (*RibbonGroup, error) {
    cr.mu.RLock()
    defer cr.mu.RUnlock()

    group, exists := cr.groups[groupID]
    if !exists {
        return nil, fmt.Errorf("group with ID %s not found", groupID)
    }

    return group, nil
}

// GetControl 获取控件
func (cr *CustomRibbon) GetControl(controlID string) (*RibbonControl, error) {
    cr.mu.RLock()
    defer cr.mu.RUnlock()

    control, exists := cr.controls[controlID]
    if !exists {
        return nil, fmt.Errorf("control with ID %s not found", controlID)
    }

    return control, nil
}

// ListTabs 列出所有选项卡
func (cr *CustomRibbon) ListTabs() []*RibbonTab {
    cr.mu.RLock()
    defer cr.mu.RUnlock()

    tabs := make([]*RibbonTab, 0, len(cr.tabs))
    for _, tab := range cr.tabs {
        tabs = append(tabs, tab)
    }

    return tabs
}

// ExportRibbon 导出功能区配置
func (cr *CustomRibbon) ExportRibbon() ([]byte, error) {
    cr.mu.RLock()
    defer cr.mu.RUnlock()

    type RibbonExport struct {
        Tabs    []*RibbonTab   `xml:"tab"`
        Config  *RibbonConfig  `xml:"config"`
        Metrics *RibbonMetrics `xml:"metrics"`
    }

    export := &RibbonExport{
        Tabs:    cr.ListTabs(),
        Config:  cr.config,
        Metrics: cr.metrics,
    }

    data, err := xml.MarshalIndent(export, "", "  ")
    if err != nil {
        return nil, fmt.Errorf("failed to export ribbon: %w", err)
    }

    return data, nil
}

// ImportRibbon 导入功能区配置
func (cr *CustomRibbon) ImportRibbon(data []byte) error {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    type RibbonImport struct {
        Tabs   []*RibbonTab  `xml:"tab"`
        Config *RibbonConfig `xml:"config"`
    }

    importData := &RibbonImport{}
    err := xml.Unmarshal(data, importData)
    if err != nil {
        return fmt.Errorf("failed to parse ribbon data: %w", err)
    }

    // 清空现有数据
    cr.tabs = make(map[string]*RibbonTab)
    cr.groups = make(map[string]*RibbonGroup)
    cr.controls = make(map[string]*RibbonControl)
    cr.callbacks = make(map[string]RibbonCallback)

    // 导入选项卡
    for _, tab := range importData.Tabs {
        cr.tabs[tab.ID] = tab

        // 导入组
        for _, group := range tab.Groups {
            cr.groups[group.ID] = group

            // 导入控件
            for _, control := range group.Controls {
                cr.controls[control.ID] = control
            }
        }
    }

    // 更新配置
    if importData.Config != nil {
        cr.config = importData.Config
        cr.isEnabled = cr.config.Enabled
    }

    // 重新计算指标
    cr.recalculateMetrics()

    cr.logger.Info("功能区导入成功，选项卡数: %d, 组数: %d, 控件数: %d", len(cr.tabs), len(cr.groups), len(cr.controls))

    return nil
}

// GetMetrics 获取指标
func (cr *CustomRibbon) GetMetrics() *RibbonMetrics {
    cr.mu.RLock()
    defer cr.mu.RUnlock()

    return cr.metrics
}

// Enable 启用功能区
func (cr *CustomRibbon) Enable() {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    cr.isEnabled = true
    cr.config.Enabled = true

    cr.logger.Info("功能区已启用")
}

// Disable 禁用功能区
func (cr *CustomRibbon) Disable() {
    cr.mu.Lock()
    defer cr.mu.Unlock()

    cr.isEnabled = false
    cr.config.Enabled = false

    cr.logger.Info("功能区已禁用")
}

// 辅助方法

// isControlTypeAllowed 检查控件类型是否允许
func (cr *CustomRibbon) isControlTypeAllowed(controlType ControlType) bool {
    for _, allowed := range cr.config.AllowedControlTypes {
        if controlType == allowed {
            return true
        }
    }
    return false
}

// getDefaultControlStyle 获取默认控件样式
func (cr *CustomRibbon) getDefaultControlStyle() *ControlStyle {
    return &ControlStyle{
        BackgroundColor: "#ffffff",
        ForegroundColor: "#000000",
        BorderColor:     "#cccccc",
        BorderWidth:     1,
        FontFamily:      "Segoe UI",
        FontSize:        12,
        FontWeight:      "normal",
        Padding:         4,
        Margin:          2,
    }
}

// recalculateMetrics 重新计算指标
func (cr *CustomRibbon) recalculateMetrics() {
    cr.metrics.TotalTabs = int64(len(cr.tabs))
    cr.metrics.TotalGroups = int64(len(cr.groups))
    cr.metrics.TotalControls = int64(len(cr.controls))

    activeCount := int64(0)
    for _, control := range cr.controls {
        if control.Enabled {
            activeCount++
        }
    }
    cr.metrics.ActiveControls = activeCount
}

// getDefaultRibbonConfig 获取默认配置
func getDefaultRibbonConfig() *RibbonConfig {
    return &RibbonConfig{
        Enabled:             true,
        MaxTabs:             10,
        MaxGroupsPerTab:     20,
        MaxControlsPerGroup: 50,
        AllowedControlTypes: []ControlType{
            ControlTypeButton,
            ControlTypeToggleButton,
            ControlTypeDropDown,
            ControlTypeComboBox,
            ControlTypeTextBox,
            ControlTypeCheckBox,
            ControlTypeRadioButton,
            ControlTypeSeparator,
            ControlTypeLabel,
        },
        Theme:            "default",
        AnimationEnabled: true,
        TooltipEnabled:   true,
        ShortcutEnabled:  true,
        AutoSave:         true,
        AutoSaveInterval: 300, // 5 minutes
    }
}
