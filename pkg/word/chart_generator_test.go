package word

import (
    "context"
    "fmt"
    "testing"
)

// TestNewChartGenerator 测试创建图表生成器
func TestNewChartGenerator(t *testing.T) {
    // 测试默认配置
    cg := NewChartGenerator(nil)
    if cg == nil {
        t.Fatal("图表生成器创建失败")
    }

    // 验证默认配置
    if cg.config.MaxCharts != 1000 {
        t.Errorf("默认最大图表数不匹配，期望: 1000, 实际: %d", cg.config.MaxCharts)
    }

    if cg.config.MaxDataPoints != 10000 {
        t.Errorf("默认最大数据点数不匹配，期望: 10000, 实际: %d", cg.config.MaxDataPoints)
    }

    if !cg.config.CacheEnabled {
        t.Error("默认缓存应该启用")
    }

    // 测试自定义配置
    customConfig := &ChartGeneratorConfig{
        MaxCharts:     500,
        MaxDataPoints: 5000,
        CacheEnabled:  false,
    }

    cg2 := NewChartGenerator(customConfig)
    if cg2.config.MaxCharts != 500 {
        t.Errorf("自定义最大图表数不匹配，期望: 500, 实际: %d", cg2.config.MaxCharts)
    }

    if cg2.config.CacheEnabled {
        t.Error("自定义配置应该禁用缓存")
    }
}

// TestCreateChart 测试创建图表
func TestCreateChart(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 测试创建基本图表
    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "测试图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    if chart.ID == "" {
        t.Error("图表ID应该自动生成")
    }

    if chart.CreatedAt == 0 {
        t.Error("创建时间应该设置")
    }

    if chart.UpdatedAt == 0 {
        t.Error("更新时间应该设置")
    }

    // 验证图表已存储
    storedChart, err := cg.GetChart(chart.ID)
    if err != nil {
        t.Fatalf("获取图表失败: %v", err)
    }

    if storedChart.Title != "测试图表" {
        t.Errorf("图表标题不匹配，期望: 测试图表, 实际: %s", storedChart.Title)
    }
}

// TestCreateChartWithInvalidType 测试创建无效类型的图表
func TestCreateChartWithInvalidType(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    chart := &ChartGeneratorChart{
        Type:  "invalid_type",
        Title: "无效图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err == nil {
        t.Error("应该拒绝无效的图表类型")
    }
}

// TestCreateChartWithExistingID 测试创建带有现有ID的图表
func TestCreateChartWithExistingID(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    chart1 := &ChartGeneratorChart{
        ID:    "test_id",
        Type:  ChartGeneratorChartTypeLine,
        Title: "图表1",
    }

    err := cg.CreateChart(ctx, chart1)
    if err != nil {
        t.Fatalf("创建第一个图表失败: %v", err)
    }

    chart2 := &ChartGeneratorChart{
        ID:    "test_id", // 相同ID
        Type:  ChartGeneratorChartTypeBar,
        Title: "图表2",
    }

    err = cg.CreateChart(ctx, chart2)
    if err != nil {
        t.Fatalf("创建第二个图表失败: %v", err)
    }

    // 验证第二个图表覆盖了第一个
    storedChart, err := cg.GetChart("test_id")
    if err != nil {
        t.Fatalf("获取图表失败: %v", err)
    }

    if storedChart.Type != ChartGeneratorChartTypeBar {
        t.Errorf("图表类型应该被覆盖，期望: %s, 实际: %s", ChartGeneratorChartTypeBar, storedChart.Type)
    }
}

// TestAddDataSource 测试添加数据源
func TestAddDataSource(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    dataSource := &ChartGeneratorDataSource{
        Name: "测试数据源",
        Type: ChartGeneratorDataSourceTypeStatic,
    }

    err := cg.AddDataSource(ctx, dataSource)
    if err != nil {
        t.Fatalf("添加数据源失败: %v", err)
    }

    if dataSource.ID == "" {
        t.Error("数据源ID应该自动生成")
    }

    // 验证数据源已存储
    storedDataSource, err := cg.GetDataSource(dataSource.ID)
    if err != nil {
        t.Fatalf("获取数据源失败: %v", err)
    }

    if storedDataSource.Name != "测试数据源" {
        t.Errorf("数据源名称不匹配，期望: 测试数据源, 实际: %s", storedDataSource.Name)
    }
}

// TestAddDataPoint 测试添加数据点
func TestAddDataPoint(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 先创建图表
    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "数据点测试图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    // 添加数据点
    dataPoint := &ChartGeneratorDataPoint{
        Label: "测试点",
        Value: 42.5,
    }

    err = cg.AddDataPoint(ctx, chart.ID, dataPoint)
    if err != nil {
        t.Fatalf("添加数据点失败: %v", err)
    }

    // 验证数据点已添加
    storedChart, err := cg.GetChart(chart.ID)
    if err != nil {
        t.Fatalf("获取图表失败: %v", err)
    }

    if len(storedChart.Data) != 1 {
        t.Errorf("数据点数量不匹配，期望: 1, 实际: %d", len(storedChart.Data))
    }

    if storedChart.Data[0].Label != "测试点" {
        t.Errorf("数据点标签不匹配，期望: 测试点, 实际: %s", storedChart.Data[0].Label)
    }

    if storedChart.Data[0].Value != 42.5 {
        t.Errorf("数据点值不匹配，期望: 42.5, 实际: %f", storedChart.Data[0].Value)
    }
}

// TestAddDataPointToNonExistentChart 测试向不存在的图表添加数据点
func TestAddDataPointToNonExistentChart(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    dataPoint := &ChartGeneratorDataPoint{
        Label: "测试点",
        Value: 42.5,
    }

    err := cg.AddDataPoint(ctx, "non_existent_id", dataPoint)
    if err == nil {
        t.Error("应该拒绝向不存在的图表添加数据点")
    }
}

// TestUpdateChartStyle 测试更新图表样式
func TestUpdateChartStyle(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 先创建图表
    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "样式测试图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    // 更新样式
    newStyle := &ChartGeneratorStyle{
        Theme:      ChartGeneratorThemeDark,
        FontFamily: "Arial",
        FontSize:   16,
    }

    err = cg.UpdateChartStyle(ctx, chart.ID, newStyle)
    if err != nil {
        t.Fatalf("更新图表样式失败: %v", err)
    }

    // 验证样式已更新
    storedChart, err := cg.GetChart(chart.ID)
    if err != nil {
        t.Fatalf("获取图表失败: %v", err)
    }

    if storedChart.Style.Theme != ChartGeneratorThemeDark {
        t.Errorf("主题应该更新，期望: %s, 实际: %s", ChartGeneratorThemeDark, storedChart.Style.Theme)
    }

    if storedChart.Style.FontFamily != "Arial" {
        t.Errorf("字体应该更新，期望: Arial, 实际: %s", storedChart.Style.FontFamily)
    }

    if storedChart.Style.FontSize != 16 {
        t.Errorf("字体大小应该更新，期望: 16, 实际: %d", storedChart.Style.FontSize)
    }
}

// TestUpdateChartOptions 测试更新图表选项
func TestUpdateChartOptions(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 先创建图表
    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "选项测试图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    // 更新选项
    newOptions := &ChartGeneratorOptions{
        Width:       800,
        Height:      600,
        Responsive:  false,
        Animation:   false,
        Interactive: false,
    }

    err = cg.UpdateChartOptions(ctx, chart.ID, newOptions)
    if err != nil {
        t.Fatalf("更新图表选项失败: %v", err)
    }

    // 验证选项已更新
    storedChart, err := cg.GetChart(chart.ID)
    if err != nil {
        t.Fatalf("获取图表失败: %v", err)
    }

    if storedChart.Options.Width != 800 {
        t.Errorf("宽度应该更新，期望: 800, 实际: %d", storedChart.Options.Width)
    }

    if storedChart.Options.Height != 600 {
        t.Errorf("高度应该更新，期望: 600, 实际: %d", storedChart.Options.Height)
    }

    if storedChart.Options.Responsive {
        t.Error("响应式应该被禁用")
    }

    if storedChart.Options.Animation {
        t.Error("动画应该被禁用")
    }

    if storedChart.Options.Interactive {
        t.Error("交互性应该被禁用")
    }
}

// TestDeleteChart 测试删除图表
func TestDeleteChart(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 先创建图表
    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "删除测试图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    // 删除图表
    err = cg.DeleteChart(ctx, chart.ID)
    if err != nil {
        t.Fatalf("删除图表失败: %v", err)
    }

    // 验证图表已被删除
    _, err = cg.GetChart(chart.ID)
    if err == nil {
        t.Error("图表应该已被删除")
    }
}

// TestDeleteNonExistentChart 测试删除不存在的图表
func TestDeleteNonExistentChart(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    err := cg.DeleteChart(ctx, "non_existent_id")
    if err == nil {
        t.Error("应该拒绝删除不存在的图表")
    }
}

// TestExportChart 测试导出图表
func TestExportChart(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 先创建图表
    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "导出测试图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    // 测试导出为SVG
    exportData, err := cg.ExportChart(ctx, chart.ID, "svg", nil)
    if err != nil {
        t.Fatalf("导出图表失败: %v", err)
    }

    if len(exportData) == 0 {
        t.Error("导出数据不应该为空")
    }

    // 测试导出为PNG
    exportData, err = cg.ExportChart(ctx, chart.ID, "png", nil)
    if err != nil {
        t.Fatalf("导出PNG图表失败: %v", err)
    }

    if len(exportData) == 0 {
        t.Error("PNG导出数据不应该为空")
    }
}

// TestExportChartWithInvalidFormat 测试导出无效格式的图表
func TestExportChartWithInvalidFormat(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 先创建图表
    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "格式测试图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    // 测试导出无效格式
    _, err = cg.ExportChart(ctx, chart.ID, "invalid_format", nil)
    if err == nil {
        t.Error("应该拒绝导出无效格式")
    }
}

// TestExportNonExistentChart 测试导出不存在的图表
func TestExportNonExistentChart(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    _, err := cg.ExportChart(ctx, "non_existent_id", "svg", nil)
    if err == nil {
        t.Error("应该拒绝导出不存在的图表")
    }
}

// TestGetStats 测试获取统计信息
func TestGetStats(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 获取初始统计
    initialStats := cg.GetStats()
    if initialStats["total_charts"].(int) != 0 {
        t.Errorf("初始图表总数应该为0，实际: %d", initialStats["total_charts"].(int))
    }

    // 创建一些图表
    charts := []*ChartGeneratorChart{
        {Type: ChartGeneratorChartTypeLine, Title: "图表1"},
        {Type: ChartGeneratorChartTypeBar, Title: "图表2"},
        {Type: ChartGeneratorChartTypePie, Title: "图表3"},
    }

    for _, chart := range charts {
        err := cg.CreateChart(ctx, chart)
        if err != nil {
            t.Fatalf("创建图表失败: %v", err)
        }
    }

    // 添加一些数据源
    dataSources := []*ChartGeneratorDataSource{
        {Name: "数据源1", Type: ChartGeneratorDataSourceTypeStatic},
        {Name: "数据源2", Type: ChartGeneratorDataSourceTypeDatabase},
    }

    for _, ds := range dataSources {
        err := cg.AddDataSource(ctx, ds)
        if err != nil {
            t.Fatalf("添加数据源失败: %v", err)
        }
    }

    // 获取更新后的统计
    updatedStats := cg.GetStats()
    if updatedStats["total_charts"].(int) != 3 {
        t.Errorf("图表总数应该为3，实际: %d", updatedStats["total_charts"].(int))
    }

    if updatedStats["total_data_sources"].(int) != 2 {
        t.Errorf("数据源总数应该为2，实际: %d", updatedStats["total_data_sources"].(int))
    }

    // 验证图表类型统计
    chartTypeCount := updatedStats["chart_type_count"].(map[ChartGeneratorChartType]int)
    if chartTypeCount[ChartGeneratorChartTypeLine] != 1 {
        t.Errorf("折线图数量应该为1，实际: %d", chartTypeCount[ChartGeneratorChartTypeLine])
    }

    if chartTypeCount[ChartGeneratorChartTypeBar] != 1 {
        t.Errorf("柱状图数量应该为1，实际: %d", chartTypeCount[ChartGeneratorChartTypeBar])
    }

    if chartTypeCount[ChartGeneratorChartTypePie] != 1 {
        t.Errorf("饼图数量应该为1，实际: %d", chartTypeCount[ChartGeneratorChartTypePie])
    }

    // 验证数据源类型统计
    dataSourceTypeCount := updatedStats["data_source_type_count"].(map[ChartGeneratorDataSourceType]int)
    if dataSourceTypeCount[ChartGeneratorDataSourceTypeStatic] != 1 {
        t.Errorf("静态数据源数量应该为1，实际: %d", dataSourceTypeCount[ChartGeneratorDataSourceTypeStatic])
    }

    if dataSourceTypeCount[ChartGeneratorDataSourceTypeDatabase] != 1 {
        t.Errorf("数据库数据源数量应该为1，实际: %d", dataSourceTypeCount[ChartGeneratorDataSourceTypeDatabase])
    }
}

// TestChartGeneratorConcurrency 测试并发安全性
func TestChartGeneratorConcurrency(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 并发创建图表
    const numGoroutines = 10
    errors := make(chan error, numGoroutines)

    for i := 0; i < numGoroutines; i++ {
        go func(id int) {
            chart := &ChartGeneratorChart{
                Type:  ChartGeneratorChartTypeLine,
                Title: fmt.Sprintf("并发图表%d", id),
            }
            err := cg.CreateChart(ctx, chart)
            errors <- err
        }(i)
    }

    // 收集错误
    for i := 0; i < numGoroutines; i++ {
        err := <-errors
        if err != nil {
            t.Errorf("并发创建图表失败: %v", err)
        }
    }

    // 验证所有图表都已创建
    stats := cg.GetStats()
    if stats["total_charts"].(int) != numGoroutines {
        t.Errorf("并发创建的图表数量不匹配，期望: %d, 实际: %d", numGoroutines, stats["total_charts"].(int))
    }
}

// TestChartGeneratorMemoryLimits 测试内存限制
func TestChartGeneratorMemoryLimits(t *testing.T) {
    // 创建限制配置
    config := &ChartGeneratorConfig{
        MaxCharts:      2,
        MaxDataPoints:  5,
        MaxDataSources: 3,
        SupportedTypes: []string{"line", "bar", "pie"},
    }

    cg := NewChartGenerator(config)
    ctx := context.Background()

    // 测试图表数量限制
    for i := 0; i < 3; i++ {
        chart := &ChartGeneratorChart{
            Type:  ChartGeneratorChartTypeLine,
            Title: fmt.Sprintf("限制测试图表%d", i),
        }
        err := cg.CreateChart(ctx, chart)
        if i < 2 {
            if err != nil {
                t.Errorf("创建第%d个图表应该成功: %v", i+1, err)
            }
        } else {
            if err == nil {
                t.Error("创建第3个图表应该失败（超过限制）")
            }
        }
    }

    // 测试数据点数量限制
    // 由于MaxCharts=2，我们需要删除一个图表来为新图表腾出空间
    // 获取第一个图表的ID
    var firstChartID string
    for id := range cg.charts {
        firstChartID = id
        break
    }

    // 删除第一个图表
    err := cg.DeleteChart(ctx, firstChartID)
    if err != nil {
        t.Fatalf("删除图表失败: %v", err)
    }

    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "数据点限制测试图表",
    }

    err = cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    for i := 0; i < 6; i++ {
        dataPoint := &ChartGeneratorDataPoint{
            Label: fmt.Sprintf("点%d", i),
            Value: float64(i),
        }
        err := cg.AddDataPoint(ctx, chart.ID, dataPoint)
        if i < 5 {
            if err != nil {
                t.Errorf("添加第%d个数据点应该成功: %v", i+1, err)
            }
        } else {
            if err == nil {
                t.Error("添加第6个数据点应该失败（超过限制）")
            }
        }
    }
}

// TestChartGeneratorTemplates 测试图表模板
func TestChartGeneratorTemplates(t *testing.T) {
    cg := NewChartGenerator(nil)

    // 测试获取默认模板
    lineTemplate, err := cg.GetTemplate("line_default")
    if err != nil {
        t.Fatalf("获取折线图模板失败: %v", err)
    }

    if lineTemplate.Type != ChartGeneratorChartTypeLine {
        t.Errorf("折线图模板类型不匹配，期望: %s, 实际: %s", ChartGeneratorChartTypeLine, lineTemplate.Type)
    }

    barTemplate, err := cg.GetTemplate("bar_default")
    if err != nil {
        t.Fatalf("获取柱状图模板失败: %v", err)
    }

    if barTemplate.Type != ChartGeneratorChartTypeBar {
        t.Errorf("柱状图模板类型不匹配，期望: %s, 实际: %s", ChartGeneratorChartTypeBar, barTemplate.Type)
    }

    pieTemplate, err := cg.GetTemplate("pie_default")
    if err != nil {
        t.Fatalf("获取饼图模板失败: %v", err)
    }

    if pieTemplate.Type != ChartGeneratorChartTypePie {
        t.Errorf("饼图模板类型不匹配，期望: %s, 实际: %s", ChartGeneratorChartTypePie, pieTemplate.Type)
    }

    // 测试获取不存在的模板
    _, err = cg.GetTemplate("non_existent_template")
    if err == nil {
        t.Error("应该拒绝获取不存在的模板")
    }
}

// TestChartGeneratorDefaultValues 测试默认值设置
func TestChartGeneratorDefaultValues(t *testing.T) {
    cg := NewChartGenerator(nil)
    ctx := context.Background()

    // 创建没有样式和选项的图表
    chart := &ChartGeneratorChart{
        Type:  ChartGeneratorChartTypeLine,
        Title: "默认值测试图表",
    }

    err := cg.CreateChart(ctx, chart)
    if err != nil {
        t.Fatalf("创建图表失败: %v", err)
    }

    // 验证默认样式已设置
    if chart.Style == nil {
        t.Fatal("默认样式应该被设置")
    }

    if chart.Style.Theme != ChartGeneratorThemeDefault {
        t.Errorf("默认主题不匹配，期望: %s, 实际: %s", ChartGeneratorThemeDefault, chart.Style.Theme)
    }

    if chart.Style.FontFamily != "Arial, sans-serif" {
        t.Errorf("默认字体不匹配，期望: Arial, sans-serif, 实际: %s", chart.Style.FontFamily)
    }

    if chart.Style.FontSize != 12 {
        t.Errorf("默认字体大小不匹配，期望: 12, 实际: %d", chart.Style.FontSize)
    }

    // 验证默认选项已设置
    if chart.Options == nil {
        t.Fatal("默认选项应该被设置")
    }

    if chart.Options.Width != 600 {
        t.Errorf("默认宽度不匹配，期望: 600, 实际: %d", chart.Options.Width)
    }

    if chart.Options.Height != 400 {
        t.Errorf("默认高度不匹配，期望: 400, 实际: %d", chart.Options.Height)
    }

    if !chart.Options.Responsive {
        t.Error("默认应该启用响应式")
    }

    if !chart.Options.Animation {
        t.Error("默认应该启用动画")
    }

    if !chart.Options.Interactive {
        t.Error("默认应该启用交互性")
    }

    // 验证数据数组已初始化
    if chart.Data == nil {
        t.Fatal("数据数组应该被初始化")
    }

    if len(chart.Data) != 0 {
        t.Errorf("初始数据数组应该为空，实际长度: %d", len(chart.Data))
    }

    // 验证元数据已初始化
    if chart.Metadata == nil {
        t.Fatal("元数据应该被初始化")
    }

    if len(chart.Metadata) != 0 {
        t.Errorf("初始元数据应该为空，实际长度: %d", len(chart.Metadata))
    }
}
