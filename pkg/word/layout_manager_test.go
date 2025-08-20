package word

import (
    "fmt"
    "testing"

    "github.com/tanqiangyes/go-word/pkg/types"
)

// TestNewLayoutManager 测试创建布局管理器
func TestNewLayoutManager(t *testing.T) {
    // 测试默认配置
    lm := NewLayoutManager()
    if lm == nil {
        t.Fatal("布局管理器创建失败")
    }

    // 验证组件已创建
    if lm.PositionManager == nil {
        t.Error("位置管理器应该被创建")
    }

    if lm.SizeManager == nil {
        t.Error("尺寸管理器应该被创建")
    }

    if lm.SpacingManager == nil {
        t.Error("间距管理器应该被创建")
    }

    if lm.LayoutAlgorithm == nil {
        t.Error("布局算法应该被创建")
    }

    if lm.Metrics == nil {
        t.Error("指标应该被创建")
    }

    if lm.Logger == nil {
        t.Error("日志器应该被创建")
    }
}

// TestNewPositionManager 测试创建位置管理器
func TestNewPositionManager(t *testing.T) {
    pm := NewPositionManager()

    if pm == nil {
        t.Fatal("位置管理器创建失败")
    }

    // 验证默认位置已加载
    if len(pm.Positions) == 0 {
        t.Error("默认位置应该被加载")
    }

    // 验证默认位置存在
    defaultPos, exists := pm.Positions["default"]
    if !exists {
        t.Error("默认位置应该存在")
    }

    if defaultPos.Type != PositionTypeFlow {
        t.Errorf("默认位置类型不匹配，期望: %s, 实际: %s", PositionTypeFlow, defaultPos.Type)
    }
}

// TestNewSizeManager 测试创建尺寸管理器
func TestNewSizeManager(t *testing.T) {
    sm := NewSizeManager()

    if sm == nil {
        t.Fatal("尺寸管理器创建失败")
    }

    // 验证默认尺寸已加载
    if len(sm.Sizes) == 0 {
        t.Error("默认尺寸应该被加载")
    }

    // 验证默认尺寸存在
    defaultSize, exists := sm.Sizes["default"]
    if !exists {
        t.Error("默认尺寸应该存在")
    }

    if defaultSize.Width != 100.0 {
        t.Errorf("默认宽度不匹配，期望: 100.0, 实际: %f", defaultSize.Width)
    }

    if defaultSize.Height != 100.0 {
        t.Errorf("默认高度不匹配，期望: 100.0, 实际: %f", defaultSize.Height)
    }
}

// TestNewSpacingManager 测试创建间距管理器
func TestNewSpacingManager(t *testing.T) {
    spm := NewSpacingManager()

    if spm == nil {
        t.Fatal("间距管理器创建失败")
    }

    // 验证默认间距已加载
    if len(spm.Margins) == 0 {
        t.Error("默认外边距应该被加载")
    }

    if len(spm.Paddings) == 0 {
        t.Error("默认内边距应该被加载")
    }

    // 验证默认外边距存在
    defaultMargin, exists := spm.Margins["default"]
    if !exists {
        t.Error("默认外边距应该存在")
    }

    if defaultMargin.Top != 0.0 {
        t.Errorf("默认上边距不匹配，期望: 0.0, 实际: %f", defaultMargin.Top)
    }
}

// TestProcessLayout 测试布局处理
func TestProcessLayout(t *testing.T) {
    lm := NewLayoutManager()

    // 测试基本布局处理 - 添加更多内容以确保有可测量的处理时间
    content := &types.DocumentContent{
        Paragraphs: make([]types.Paragraph, 100), // 增加段落数量
        Tables: make([]types.Table, 10),          // 增加表格数量
    }

    // 填充段落内容
    for i := range content.Paragraphs {
        content.Paragraphs[i] = types.Paragraph{
            Text: fmt.Sprintf("测试段落%d", i+1),
        }
    }

    // 填充表格内容
    for i := range content.Tables {
        content.Tables[i] = types.Table{
            Rows: []types.TableRow{},
        }
    }

    err := lm.ProcessLayout(content)
    if err != nil {
        t.Fatalf("布局处理失败: %v", err)
    }

    // 验证指标已更新
    if lm.Metrics.ElementsPositioned == 0 {
        t.Error("元素定位指标应该被更新")
    }

    if lm.Metrics.ProcessingTime == 0 {
        t.Error("处理时间应该被记录")
    }
}

// TestProcessLayoutWithEmptyContent 测试空内容布局处理
func TestProcessLayoutWithEmptyContent(t *testing.T) {
    lm := NewLayoutManager()

    // 测试空内容
    content := &types.DocumentContent{
        Paragraphs: []types.Paragraph{},
        Tables:     []types.Table{},
    }

    err := lm.ProcessLayout(content)
    if err != nil {
        t.Fatalf("空内容布局处理失败: %v", err)
    }

    // 验证指标
    if lm.Metrics.ElementsPositioned != 0 {
        t.Error("空内容不应该有元素被定位")
    }
}

// TestLayoutManagerMetrics 测试指标收集
func TestLayoutManagerMetrics(t *testing.T) {
    lm := NewLayoutManager()

    // 处理更多内容以确保有可测量的处理时间
    content := &types.DocumentContent{
        Paragraphs: make([]types.Paragraph, 50), // 增加段落数量
        Tables: make([]types.Table, 5),          // 增加表格数量
    }

    // 填充段落内容
    for i := range content.Paragraphs {
        content.Paragraphs[i] = types.Paragraph{
            Text: fmt.Sprintf("测试段落%d内容", i+1),
        }
    }

    // 填充表格内容
    for i := range content.Tables {
        content.Tables[i] = types.Table{
            Rows: []types.TableRow{},
        }
    }

    err := lm.ProcessLayout(content)
    if err != nil {
        t.Fatalf("布局处理失败: %v", err)
    }

    // 验证指标
    expectedElements := int64(len(content.Paragraphs) + len(content.Tables))
    if lm.Metrics.ElementsPositioned != expectedElements {
        t.Errorf("元素定位数量不匹配，期望: %d, 实际: %d", expectedElements, lm.Metrics.ElementsPositioned)
    }

    if lm.Metrics.ProcessingTime <= 0 {
        t.Error("处理时间应该大于0")
    }

    if lm.Metrics.Errors != 0 {
        t.Errorf("错误数量应该为0，实际: %d", lm.Metrics.Errors)
    }
}

// TestLayoutManagerErrorHandling 测试错误处理
func TestLayoutManagerErrorHandling(t *testing.T) {
    lm := NewLayoutManager()

    // 测试nil内容
    err := lm.ProcessLayout(nil)
    if err == nil {
        t.Error("nil内容应该被拒绝")
    }
}
