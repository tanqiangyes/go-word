package wordprocessingml

import (
	"context"
	"fmt"
	"sync"

	"github.com/tanqiangyes/go-word/pkg/utils"
)

// ChartGenerator 图表生成器
type ChartGenerator struct {
	charts       map[string]*ChartGeneratorChart
	dataSources  map[string]*ChartGeneratorDataSource
	templates    map[string]*ChartGeneratorTemplate
	mu           sync.RWMutex
	logger       *utils.Logger
	config       *ChartGeneratorConfig
}

// ChartGeneratorChart 图表
type ChartGeneratorChart struct {
	ID          string                    `json:"id"`
	Type        ChartGeneratorChartType   `json:"type"`
	Title       string                    `json:"title"`
	DataSource  string                    `json:"data_source"`
	Data        []*ChartGeneratorDataPoint `json:"data"`
	Style       *ChartGeneratorStyle      `json:"style"`
	Options     *ChartGeneratorOptions    `json:"options"`
	Metadata    map[string]interface{}    `json:"metadata"`
	CreatedAt   int64                     `json:"created_at"`
	UpdatedAt   int64                     `json:"updated_at"`
}

// ChartGeneratorDataPoint 数据点
type ChartGeneratorDataPoint struct {
	Label       string      `json:"label"`
	Value       float64     `json:"value"`
	Category    string      `json:"category"`
	Series      string      `json:"series"`
	Color       string      `json:"color"`
	Metadata    map[string]interface{} `json:"metadata"`
}

// ChartGeneratorDataSource 数据源
type ChartGeneratorDataSource struct {
	ID          string                    `json:"id"`
	Name        string                    `json:"name"`
	Type        ChartGeneratorDataSourceType `json:"type"`
	Connection  string                    `json:"connection"`
	Query       string                    `json:"query"`
	Parameters  map[string]interface{}    `json:"parameters"`
	CacheEnabled bool                     `json:"cache_enabled"`
	CacheTTL    int64                     `json:"cache_ttl"`
}

// ChartGeneratorStyle 图表样式
type ChartGeneratorStyle struct {
	Theme       ChartGeneratorTheme       `json:"theme"`
	Colors      []string                 `json:"colors"`
	FontFamily  string                   `json:"font_family"`
	FontSize    int                      `json:"font_size"`
	Background  string                   `json:"background"`
	Border      *ChartGeneratorBorder     `json:"border"`
	Legend      *ChartGeneratorLegend     `json:"legend"`
	Grid        *ChartGeneratorGrid       `json:"grid"`
}

// ChartGeneratorOptions 图表选项
type ChartGeneratorOptions struct {
	Width       int                       `json:"width"`
	Height      int                       `json:"height"`
	Responsive  bool                      `json:"responsive"`
	Animation   bool                      `json:"animation"`
	Interactive bool                      `json:"interactive"`
	Tooltips    bool                      `json:"tooltips"`
	Legend      bool                      `json:"legend"`
	Grid        bool                      `json:"grid"`
	Axis        *ChartGeneratorAxis       `json:"axis"`
}

// ChartGeneratorTemplate 图表模板
type ChartGeneratorTemplate struct {
	ID          string                    `json:"id"`
	Name        string                    `json:"name"`
	Type        ChartGeneratorChartType   `json:"type"`
	Style       *ChartGeneratorStyle      `json:"style"`
	Options     *ChartGeneratorOptions    `json:"options"`
	Preview     string                    `json:"preview"`
	Category    string                    `json:"category"`
}

// ChartGeneratorConfig 配置
type ChartGeneratorConfig struct {
	MaxCharts           int           `json:"max_charts"`
	MaxDataPoints       int           `json:"max_data_points"`
	MaxDataSources      int           `json:"max_data_sources"`
	SupportedTypes      []string      `json:"supported_types"`
	DefaultTheme        string        `json:"default_theme"`
	CacheEnabled        bool          `json:"cache_enabled"`
	CacheSize           int           `json:"cache_size"`
	AnimationDuration   int           `json:"animation_duration"`
	ExportFormats       []string      `json:"export_formats"`
}

// 常量定义
const (
	// 图表类型
	ChartGeneratorChartTypeLine      ChartGeneratorChartType = "line"
	ChartGeneratorChartTypeBar       ChartGeneratorChartType = "bar"
	ChartGeneratorChartTypePie       ChartGeneratorChartType = "pie"
	ChartGeneratorChartTypeDoughnut  ChartGeneratorChartType = "doughnut"
	ChartGeneratorChartTypeArea      ChartGeneratorChartType = "area"
	ChartGeneratorChartTypeScatter   ChartGeneratorChartType = "scatter"
	ChartGeneratorChartTypeBubble    ChartGeneratorChartType = "bubble"
	ChartGeneratorChartTypeRadar     ChartGeneratorChartType = "radar"
	ChartGeneratorChartTypePolar     ChartGeneratorChartType = "polar"
	ChartGeneratorChartTypeFunnel    ChartGeneratorChartType = "funnel"
	ChartGeneratorChartTypePyramid   ChartGeneratorChartType = "pyramid"
	ChartGeneratorChartTypeGauge     ChartGeneratorChartType = "gauge"

	// 数据源类型
	ChartGeneratorDataSourceTypeStatic    ChartGeneratorDataSourceType = "static"
	ChartGeneratorDataSourceTypeDatabase  ChartGeneratorDataSourceType = "database"
	ChartGeneratorDataSourceTypeAPI       ChartGeneratorDataSourceType = "api"
	ChartGeneratorDataSourceTypeFile      ChartGeneratorDataSourceType = "file"
	ChartGeneratorDataSourceTypeStream    ChartGeneratorDataSourceType = "stream"

	// 主题
	ChartGeneratorThemeDefault   ChartGeneratorTheme = "default"
	ChartGeneratorThemeDark      ChartGeneratorTheme = "dark"
	ChartGeneratorThemeLight     ChartGeneratorTheme = "light"
	ChartGeneratorThemeColorful  ChartGeneratorTheme = "colorful"
	ChartGeneratorThemeMinimal   ChartGeneratorTheme = "minimal"
	ChartGeneratorThemeCorporate ChartGeneratorTheme = "corporate"

	// 边框样式
	ChartGeneratorBorderStyleNone   ChartGeneratorBorderStyle = "none"
	ChartGeneratorBorderStyleSolid  ChartGeneratorBorderStyle = "solid"
	ChartGeneratorBorderStyleDashed ChartGeneratorBorderStyle = "dashed"
	ChartGeneratorBorderStyleDotted ChartGeneratorBorderStyle = "dotted"

	// 图例位置
	ChartGeneratorLegendPositionTop    ChartGeneratorLegendPosition = "top"
	ChartGeneratorLegendPositionBottom ChartGeneratorLegendPosition = "bottom"
	ChartGeneratorLegendPositionLeft   ChartGeneratorLegendPosition = "left"
	ChartGeneratorLegendPositionRight  ChartGeneratorLegendPosition = "right"
)

// 类型定义
type ChartGeneratorChartType string
type ChartGeneratorDataSourceType string
type ChartGeneratorTheme string
type ChartGeneratorBorderStyle string
type ChartGeneratorLegendPosition string

// ChartGeneratorBorder 边框
type ChartGeneratorBorder struct {
	Style     ChartGeneratorBorderStyle `json:"style"`
	Width     int                      `json:"width"`
	Color     string                   `json:"color"`
	Radius    int                      `json:"radius"`
}

// ChartGeneratorLegend 图例
type ChartGeneratorLegend struct {
	Enabled   bool                      `json:"enabled"`
	Position  ChartGeneratorLegendPosition `json:"position"`
	FontSize  int                       `json:"font_size"`
	FontColor string                    `json:"font_color"`
}

// ChartGeneratorGrid 网格
type ChartGeneratorGrid struct {
	Enabled   bool   `json:"enabled"`
	Color     string `json:"color"`
	Width     int    `json:"width"`
	Opacity   float64 `json:"opacity"`
}

// ChartGeneratorAxis 坐标轴
type ChartGeneratorAxis struct {
	XEnabled  bool   `json:"x_enabled"`
	YEnabled  bool   `json:"y_enabled"`
	XLabel    string `json:"x_label"`
	YLabel    string `json:"y_label"`
	XMin      *float64 `json:"x_min"`
	XMax      *float64 `json:"x_max"`
	YMin      *float64 `json:"y_min"`
	YMax      *float64 `json:"y_max"`
}

// NewChartGenerator 创建新的图表生成器
func NewChartGenerator(config *ChartGeneratorConfig) *ChartGenerator {
	if config == nil {
		config = &ChartGeneratorConfig{
			MaxCharts:         1000,
			MaxDataPoints:     10000,
			MaxDataSources:    100,
			SupportedTypes:    []string{"line", "bar", "pie", "doughnut", "area", "scatter", "bubble", "radar", "polar", "funnel", "pyramid", "gauge"},
			DefaultTheme:      "default",
			CacheEnabled:      true,
			CacheSize:         100,
			AnimationDuration: 1000,
			ExportFormats:     []string{"png", "jpg", "svg", "pdf"},
		}
	}

	cg := &ChartGenerator{
		charts:      make(map[string]*ChartGeneratorChart),
		dataSources: make(map[string]*ChartGeneratorDataSource),
		templates:   make(map[string]*ChartGeneratorTemplate),
		config:      config,
		logger:      utils.NewLogger(utils.LogLevelInfo, nil),
	}

	// 初始化默认模板
	cg.initializeTemplates()

	return cg
}

// CreateChart 创建图表
func (cg *ChartGenerator) CreateChart(ctx context.Context, chart *ChartGeneratorChart) error {
	cg.mu.Lock()
	defer cg.mu.Unlock()

	// 检查图表数量限制
	if len(cg.charts) >= cg.config.MaxCharts {
		return utils.NewStructuredDocumentError(utils.ErrResourceExhausted, "已达到最大图表数量限制")
	}

	// 验证图表类型
	if !cg.isChartTypeSupported(chart.Type) {
		return utils.NewStructuredDocumentError(utils.ErrFormatUnsupported, fmt.Sprintf("不支持的图表类型: %s", chart.Type))
	}

	// 生成ID
	if chart.ID == "" {
		chart.ID = utils.GenerateID()
	}

	// 设置默认值
	if chart.Style == nil {
		chart.Style = cg.getDefaultStyle()
	}
	if chart.Options == nil {
		chart.Options = cg.getDefaultOptions()
	}
	if chart.Data == nil {
		chart.Data = make([]*ChartGeneratorDataPoint, 0)
	}
	if chart.Metadata == nil {
		chart.Metadata = make(map[string]interface{})
	}

	chart.CreatedAt = utils.GetCurrentTimestamp()
	chart.UpdatedAt = utils.GetCurrentTimestamp()

	// 存储图表
	cg.charts[chart.ID] = chart

	cg.logger.Info("图表已创建", map[string]interface{}{
		"chart_id": chart.ID,
		"title":    chart.Title,
		"type":     chart.Type,
	})

	return nil
}

// AddDataSource 添加数据源
func (cg *ChartGenerator) AddDataSource(ctx context.Context, dataSource *ChartGeneratorDataSource) error {
	cg.mu.Lock()
	defer cg.mu.Unlock()

	// 检查数据源数量限制
	if len(cg.dataSources) >= cg.config.MaxDataSources {
		return utils.NewStructuredDocumentError(utils.ErrResourceExhausted, "已达到最大数据源数量限制")
	}

	// 生成ID
	if dataSource.ID == "" {
		dataSource.ID = utils.GenerateID()
	}

	// 存储数据源
	cg.dataSources[dataSource.ID] = dataSource

	cg.logger.Info("数据源已添加", map[string]interface{}{
		"data_source_id": dataSource.ID,
		"name":           dataSource.Name,
		"type":           dataSource.Type,
	})

	return nil
}

// AddDataPoint 添加数据点
func (cg *ChartGenerator) AddDataPoint(ctx context.Context, chartID string, dataPoint *ChartGeneratorDataPoint) error {
	cg.mu.Lock()
	defer cg.mu.Unlock()

	chart, exists := cg.charts[chartID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图表不存在")
	}

	// 检查数据点数量限制
	if len(chart.Data) >= cg.config.MaxDataPoints {
		return utils.NewStructuredDocumentError(utils.ErrResourceExhausted, "已达到最大数据点数量限制")
	}

	// 添加数据点
	chart.Data = append(chart.Data, dataPoint)
	chart.UpdatedAt = utils.GetCurrentTimestamp()

	cg.logger.Info("数据点已添加", map[string]interface{}{
		"chart_id": chartID,
		"label":    dataPoint.Label,
		"value":    dataPoint.Value,
	})

	return nil
}

// UpdateChartStyle 更新图表样式
func (cg *ChartGenerator) UpdateChartStyle(ctx context.Context, chartID string, style *ChartGeneratorStyle) error {
	cg.mu.Lock()
	defer cg.mu.Unlock()

	chart, exists := cg.charts[chartID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图表不存在")
	}

	// 更新样式
	chart.Style = style
	chart.UpdatedAt = utils.GetCurrentTimestamp()

	cg.logger.Info("图表样式已更新", map[string]interface{}{
		"chart_id": chartID,
		"theme":    style.Theme,
	})

	return nil
}

// UpdateChartOptions 更新图表选项
func (cg *ChartGenerator) UpdateChartOptions(ctx context.Context, chartID string, options *ChartGeneratorOptions) error {
	cg.mu.Lock()
	defer cg.mu.Unlock()

	chart, exists := cg.charts[chartID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图表不存在")
	}

	// 更新选项
	chart.Options = options
	chart.UpdatedAt = utils.GetCurrentTimestamp()

	cg.logger.Info("图表选项已更新", map[string]interface{}{
		"chart_id": chartID,
		"width":    options.Width,
		"height":   options.Height,
	})

	return nil
}

// GetChart 获取图表
func (cg *ChartGenerator) GetChart(chartID string) (*ChartGeneratorChart, error) {
	cg.mu.RLock()
	defer cg.mu.RUnlock()

	chart, exists := cg.charts[chartID]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图表不存在")
	}

	return chart, nil
}

// GetDataSource 获取数据源
func (cg *ChartGenerator) GetDataSource(dataSourceID string) (*ChartGeneratorDataSource, error) {
	cg.mu.RLock()
	defer cg.mu.RUnlock()

	dataSource, exists := cg.dataSources[dataSourceID]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "数据源不存在")
	}

	return dataSource, nil
}

// GetTemplate 获取模板
func (cg *ChartGenerator) GetTemplate(templateID string) (*ChartGeneratorTemplate, error) {
	cg.mu.RLock()
	defer cg.mu.RUnlock()

	template, exists := cg.templates[templateID]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "模板不存在")
	}

	return template, nil
}

// DeleteChart 删除图表
func (cg *ChartGenerator) DeleteChart(ctx context.Context, chartID string) error {
	cg.mu.Lock()
	defer cg.mu.Unlock()

	chart, exists := cg.charts[chartID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图表不存在")
	}

	// 删除图表
	delete(cg.charts, chartID)

	cg.logger.Info("图表已删除", map[string]interface{}{
		"chart_id": chartID,
		"title":    chart.Title,
	})

	return nil
}

// ExportChart 导出图表
func (cg *ChartGenerator) ExportChart(ctx context.Context, chartID string, format string, options map[string]interface{}) ([]byte, error) {
	chart, err := cg.GetChart(chartID)
	if err != nil {
		return nil, err
	}

	// 检查导出格式是否支持
	if !cg.isExportFormatSupported(format) {
		return nil, utils.NewStructuredDocumentError(utils.ErrFormatUnsupported, fmt.Sprintf("不支持的导出格式: %s", format))
	}

	// 这里应该实现真正的图表导出
	// 为了简化，我们返回模拟数据
	exportData := []byte(fmt.Sprintf("Chart: %s, Type: %s, Format: %s", chart.Title, chart.Type, format))

	cg.logger.Info("图表已导出", map[string]interface{}{
		"chart_id": chartID,
		"format":   format,
		"size":     len(exportData),
	})

	return exportData, nil
}

// GetStats 获取统计信息
func (cg *ChartGenerator) GetStats() map[string]interface{} {
	cg.mu.RLock()
	defer cg.mu.RUnlock()

	stats := map[string]interface{}{
		"total_charts":      len(cg.charts),
		"total_data_sources": len(cg.dataSources),
		"total_templates":   len(cg.templates),
		"supported_types":   cg.config.SupportedTypes,
		"export_formats":    cg.config.ExportFormats,
	}

	// 按图表类型统计
	chartTypeCount := make(map[ChartGeneratorChartType]int)
	for _, chart := range cg.charts {
		chartTypeCount[chart.Type]++
	}
	stats["chart_type_count"] = chartTypeCount

	// 按数据源类型统计
	dataSourceTypeCount := make(map[ChartGeneratorDataSourceType]int)
	for _, dataSource := range cg.dataSources {
		dataSourceTypeCount[dataSource.Type]++
	}
	stats["data_source_type_count"] = dataSourceTypeCount

	return stats
}

// 辅助方法

// initializeTemplates 初始化默认模板
func (cg *ChartGenerator) initializeTemplates() {
	templates := map[string]*ChartGeneratorTemplate{
		"line_default": {
			ID:       "line_default",
			Name:     "默认折线图",
			Type:     ChartGeneratorChartTypeLine,
			Style:    cg.getDefaultStyle(),
			Options:  cg.getDefaultOptions(),
			Category: "line",
		},
		"bar_default": {
			ID:       "bar_default",
			Name:     "默认柱状图",
			Type:     ChartGeneratorChartTypeBar,
			Style:    cg.getDefaultStyle(),
			Options:  cg.getDefaultOptions(),
			Category: "bar",
		},
		"pie_default": {
			ID:       "pie_default",
			Name:     "默认饼图",
			Type:     ChartGeneratorChartTypePie,
			Style:    cg.getDefaultStyle(),
			Options:  cg.getDefaultOptions(),
			Category: "pie",
		},
	}

	for key, template := range templates {
		cg.templates[key] = template
	}
}

// isChartTypeSupported 检查图表类型是否支持
func (cg *ChartGenerator) isChartTypeSupported(chartType ChartGeneratorChartType) bool {
	for _, supported := range cg.config.SupportedTypes {
		if supported == string(chartType) {
			return true
		}
	}
	return false
}

// isExportFormatSupported 检查导出格式是否支持
func (cg *ChartGenerator) isExportFormatSupported(format string) bool {
	for _, supported := range cg.config.ExportFormats {
		if supported == format {
			return true
		}
	}
	return false
}

// getDefaultStyle 获取默认样式
func (cg *ChartGenerator) getDefaultStyle() *ChartGeneratorStyle {
	return &ChartGeneratorStyle{
		Theme:      ChartGeneratorThemeDefault,
		Colors:     []string{"#007bff", "#28a745", "#dc3545", "#ffc107", "#17a2b8"},
		FontFamily: "Arial, sans-serif",
		FontSize:   12,
		Background: "#ffffff",
		Border: &ChartGeneratorBorder{
			Style:  ChartGeneratorBorderStyleSolid,
			Width:  1,
			Color:  "#e9ecef",
			Radius: 4,
		},
		Legend: &ChartGeneratorLegend{
			Enabled:   true,
			Position:  ChartGeneratorLegendPositionBottom,
			FontSize:  12,
			FontColor: "#6c757d",
		},
		Grid: &ChartGeneratorGrid{
			Enabled: true,
			Color:   "#f8f9fa",
			Width:   1,
			Opacity: 0.5,
		},
	}
}

// getDefaultOptions 获取默认选项
func (cg *ChartGenerator) getDefaultOptions() *ChartGeneratorOptions {
	return &ChartGeneratorOptions{
		Width:       600,
		Height:      400,
		Responsive:  true,
		Animation:   true,
		Interactive: true,
		Tooltips:    true,
		Legend:      true,
		Grid:        true,
		Axis: &ChartGeneratorAxis{
			XEnabled: true,
			YEnabled: true,
			XLabel:   "X轴",
			YLabel:   "Y轴",
		},
	}
}
