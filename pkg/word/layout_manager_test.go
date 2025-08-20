package word

import (
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
    if lm.positionManager == nil {
        t.Error("位置管理器应该被创建")
    }

    if lm.sizeManager == nil {
        t.Error("尺寸管理器应该被创建")
    }

    if lm.spacingManager == nil {
        t.Error("间距管理器应该被创建")
    }

    if lm.layoutAlgorithm == nil {
        t.Error("布局算法应该被创建")
    }

    if lm.metrics == nil {
        t.Error("指标应该被创建")
    }

    if lm.logger == nil {
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
    if len(pm.positions) == 0 {
        t.Error("默认位置应该被加载")
    }

    // 验证默认位置存在
    defaultPos, exists := pm.positions["default"]
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
    if len(sm.sizes) == 0 {
        t.Error("默认尺寸应该被加载")
    }

    // 验证默认尺寸存在
    defaultSize, exists := sm.sizes["default"]
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
    if len(spm.margins) == 0 {
        t.Error("默认外边距应该被加载")
    }

    if len(spm.paddings) == 0 {
        t.Error("默认内边距应该被加载")
    }

    // 验证默认外边距存在
    defaultMargin, exists := spm.margins["default"]
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

    // 测试基本布局处理
    content := &types.DocumentContent{
        Paragraphs: []types.Paragraph{
            {
                Text: "测试段落1",
            },
            {
                Text: "测试段落2",
            },
        },
        Tables: []types.Table{
            {
                Rows: []types.TableRow{},
            },
        },
    }

    err := lm.ProcessLayout(content)
    if err != nil {
        t.Fatalf("布局处理失败: %v", err)
    }

    // 验证指标已更新
    if lm.metrics.ElementsPositioned == 0 {
        t.Error("元素定位指标应该被更新")
    }

    if lm.metrics.ProcessingTime == 0 {
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
    if lm.metrics.ElementsPositioned != 0 {
        t.Error("空内容不应该有元素被定位")
    }
}

// TestLayoutManagerMetrics 测试指标收集
func TestLayoutManagerMetrics(t *testing.T) {
    lm := NewLayoutManager()

    // 处理一些内容
    content := &types.DocumentContent{
        Paragraphs: []types.Paragraph{
            {Text: "段落1"},
            {Text: "段落2"},
        },
        Tables: []types.Table{
            {Rows: []types.TableRow{}},
        },
    }

    err := lm.ProcessLayout(content)
    if err != nil {
        t.Fatalf("布局处理失败: %v", err)
    }

    // 验证指标
    if lm.metrics.ElementsPositioned != 3 {
        t.Errorf("元素定位数量不匹配，期望: 3, 实际: %d", lm.metrics.ElementsPositioned)
    }

    if lm.metrics.ProcessingTime <= 0 {
        t.Error("处理时间应该大于0")
    }

    if lm.metrics.Errors != 0 {
        t.Errorf("错误数量应该为0，实际: %d", lm.metrics.Errors)
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
