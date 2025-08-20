package word

import (
    "fmt"
    "time"

    "github.com/tanqiangyes/go-word/pkg/types"
    "github.com/tanqiangyes/go-word/pkg/utils"
)

// StyleMerger represents a style merger
type StyleMerger struct {
    logger *utils.Logger
    styles map[string]*types.Style
}

// NewStyleMerger creates a new style merger
func NewStyleMerger() *StyleMerger {
    return &StyleMerger{
        logger: utils.NewLogger(utils.LogLevelInfo, nil),
        styles: make(map[string]*types.Style),
    }
}

// MergeStyles 合并样式
func (sm *StyleMerger) MergeStyles(conflict *types.StyleConflict) error {
    sm.logger.Info("开始合并样式，样式ID: %s, 冲突类型: %s, 优先级: %d", conflict.StyleID, conflict.Type, conflict.Priority)

    // 获取冲突的样式
    originalStyle := conflict.OriginalStyle
    if originalStyle == nil {
        return fmt.Errorf("原始样式未找到: %s", conflict.StyleID)
    }

    newStyle := conflict.NewStyle
    if newStyle == nil {
        return fmt.Errorf("新样式为空")
    }

    // 根据冲突类型选择合并策略
    switch conflict.Type {
    case types.StyleConflictTypeProperty:
        return sm.mergePropertyConflicts(originalStyle, newStyle, conflict)
    case types.StyleConflictTypeInheritance:
        return sm.mergePropertyConflicts(originalStyle, newStyle, conflict)
    case types.StyleConflictTypePriority:
        return sm.mergePropertyConflicts(originalStyle, newStyle, conflict)
    case types.StyleConflictTypeFormat:
        return sm.mergePropertyConflicts(originalStyle, newStyle, conflict)
    default:
        return sm.mergeDefaultStrategy(originalStyle, newStyle, conflict)
    }
}

// mergePropertyConflicts 合并属性冲突
func (sm *StyleMerger) mergePropertyConflicts(original, new *types.Style, conflict *types.StyleConflict) error {
    sm.logger.Info("合并属性冲突，样式ID: %s, 冲突数量: %d", original.ID, len(conflict.ConflictingProperties))

    // 创建合并后的样式
    mergedStyle := &types.Style{
        ID:         original.ID,
        Name:       original.Name,
        Type:       original.Type,
        BasedOn:    original.BasedOn,
        Next:       original.Next,
        Properties: &types.StyleProperties{},
        CreatedAt:  original.CreatedAt,
        UpdatedAt:  &time.Time{},
    }

    // 合并基础属性
    mergedStyle.Description = sm.mergeStringProperty(original.Description, new.Description, "description")
    mergedStyle.Category = sm.mergeStringProperty(original.Category, new.Category, "category")
    mergedStyle.Aliases = sm.mergeStringSlice(original.Aliases, new.Aliases)

    // 合并样式属性
    if original.Properties != nil && new.Properties != nil {
        mergedStyle.Properties = sm.mergeStyleProperties(original.Properties, new.Properties, conflict)
    } else if original.Properties != nil {
        mergedStyle.Properties = original.Properties
    } else if new.Properties != nil {
        mergedStyle.Properties = new.Properties
    }

    // 更新冲突结果
    conflict.ResolvedStyle = mergedStyle
    conflict.ResolvedAt = &time.Time{}
    conflict.Status = types.StyleConflictStatusResolved

    sm.logger.Info("属性冲突合并完成，样式ID: %s, 合并属性数: %d", mergedStyle.ID, len(conflict.ConflictingProperties))

    return nil
}

// mergeInheritanceConflicts 合并继承冲突
func (sm *StyleMerger) mergeInheritanceConflicts(original, new *types.Style, conflict *types.StyleConflict) error {
    sm.logger.Info("合并继承冲突，样式ID: %s, 原始继承: %s, 新继承: %s", original.ID, original.BasedOn, new.BasedOn)

    // 分析继承链
    originalChain := sm.getInheritanceChain(original)
    newChain := sm.getInheritanceChain(new)

    // 选择最优的继承链
    optimalChain := sm.selectOptimalInheritanceChain(originalChain, newChain, conflict.Priority)

    // 更新样式
    mergedStyle := original.Clone()
    mergedStyle.BasedOn = optimalChain[0]
    mergedStyle.UpdatedAt = &time.Time{}

    // 更新冲突结果
    conflict.ResolvedStyle = mergedStyle
    conflict.ResolvedAt = &time.Time{}
    conflict.Status = types.StyleConflictStatusResolved

    sm.logger.Info("继承冲突合并完成，样式ID: %s, 新继承: %s, 链长度: %d", mergedStyle.ID, mergedStyle.BasedOn, len(optimalChain))

    return nil
}

// mergePriorityConflicts 合并优先级冲突
func (sm *StyleMerger) mergePriorityConflicts(original, new *types.Style, conflict *types.StyleConflict) error {
    sm.logger.Info("合并优先级冲突，样式ID: %s, 原始优先级: %d, 新优先级: %d", original.ID, conflict.OriginalPriority, conflict.NewPriority)

    // 根据优先级选择样式
    var selectedStyle *types.Style
    if conflict.NewPriority > conflict.OriginalPriority {
        selectedStyle = new
        sm.logger.Info("选择新样式（更高优先级），样式ID: %s", new.ID)
    } else {
        selectedStyle = original
        sm.logger.Info("保留原始样式（更高或相等优先级），样式ID: %s", original.ID)
    }

    // 更新冲突结果
    conflict.ResolvedStyle = selectedStyle
    conflict.ResolvedAt = &time.Time{}
    conflict.Status = types.StyleConflictStatusResolved

    return nil
}

// mergeFormatConflicts 合并格式冲突
func (sm *StyleMerger) mergeFormatConflicts(original, new *types.Style, conflict *types.StyleConflict) error {
    sm.logger.Info("合并格式冲突，样式ID: %s, 冲突数量: %d", original.ID, len(conflict.ConflictingProperties))

    // 创建合并后的样式
    mergedStyle := original.Clone()
    mergedStyle.UpdatedAt = &time.Time{}

    // 合并格式属性
    if original.Properties != nil && new.Properties != nil {
        mergedStyle.Properties = sm.mergeFormatProperties(original.Properties, new.Properties, conflict)
    }

    // 更新冲突结果
    conflict.ResolvedStyle = mergedStyle
    conflict.ResolvedAt = &time.Time{}
    conflict.Status = types.StyleConflictStatusResolved

    sm.logger.Info("格式冲突合并完成，样式ID: %s, 合并格式数: %d", mergedStyle.ID, len(conflict.ConflictingProperties))

    return nil
}

// mergeDefaultStrategy 默认合并策略
func (sm *StyleMerger) mergeDefaultStrategy(original, new *types.Style, conflict *types.StyleConflict) error {
    sm.logger.Info("使用默认合并策略，样式ID: %s, 策略: %s", original.ID, "conservative")

    // 保守策略：保留原始样式，只添加新样式中不冲突的属性
    mergedStyle := original.Clone()
    mergedStyle.UpdatedAt = &time.Time{}

    // 合并非冲突属性
    if new.Properties != nil {
        if mergedStyle.Properties == nil {
            mergedStyle.Properties = &types.StyleProperties{}
        }

        // 安全地合并属性
        sm.safeMergeProperties(mergedStyle.Properties, new.Properties)
    }

    // 更新冲突结果
    conflict.ResolvedStyle = mergedStyle
    conflict.ResolvedAt = &time.Time{}
    conflict.Status = types.StyleConflictStatusResolved

    return nil
}

// 辅助方法
func (sm *StyleMerger) mergeStringProperty(original, new, propertyName string) string {
    if new != "" && new != original {
        sm.logger.Debug("合并字符串属性，属性: %s, 原始: %s, 新: %s", propertyName, original, new)
        return new
    }
    return original
}

func (sm *StyleMerger) mergeStringSlice(original, new []string) []string {
    if len(new) > 0 {
        // 合并并去重
        merged := make(map[string]bool)
        for _, item := range original {
            merged[item] = true
        }
        for _, item := range new {
            merged[item] = true
        }

        result := make([]string, 0, len(merged))
        for item := range merged {
            result = append(result, item)
        }
        return result
    }
    return original
}

// mergeStyleProperties 合并样式属性
func (sm *StyleMerger) mergeStyleProperties(original, new *types.StyleProperties, conflict *types.StyleConflict) *types.StyleProperties {
    merged := &types.StyleProperties{}

    // 合并字体属性
    if original.FontName != "" && new.FontName != "" {
        merged.FontName = sm.mergeStringProperty(original.FontName, new.FontName, "fontName")
    } else if original.FontName != "" {
        merged.FontName = original.FontName
    } else if new.FontName != "" {
        merged.FontName = new.FontName
    }

    // 合并字体大小
    if original.FontSize > 0 && new.FontSize > 0 {
        merged.FontSize = sm.mergeIntProperty(original.FontSize, new.FontSize, "fontSize")
    } else if original.FontSize > 0 {
        merged.FontSize = original.FontSize
    } else if new.FontSize > 0 {
        merged.FontSize = new.FontSize
    }

    // 合并布尔属性
    merged.Bold = sm.mergeBoolProperty(original.Bold, new.Bold, "bold")
    merged.Italic = sm.mergeBoolProperty(original.Italic, new.Italic, "italic")
    merged.Underline = sm.mergeBoolProperty(original.Underline, new.Underline, "underline")

    // 合并对齐方式
    if original.Alignment != "" && new.Alignment != "" {
        merged.Alignment = sm.mergeStringProperty(original.Alignment, new.Alignment, "alignment")
    } else if original.Alignment != "" {
        merged.Alignment = original.Alignment
    } else if new.Alignment != "" {
        merged.Alignment = new.Alignment
    }

    return merged
}

// mergeFontProperties 合并字体属性
func (sm *StyleMerger) mergeFontProperties(original, new *types.Font) *types.Font {
    if original == nil && new == nil {
        return nil
    }

    if original == nil {
        return new
    }

    if new == nil {
        return original
    }

    // 创建合并后的字体
    merged := &types.Font{}

    // 合并Ascii字段
    if original.Ascii != "" && new.Ascii != "" {
        merged.Ascii = sm.mergeStringProperty(original.Ascii, new.Ascii, "fontAscii")
    } else if original.Ascii != "" {
        merged.Ascii = original.Ascii
    } else if new.Ascii != "" {
        merged.Ascii = new.Ascii
    }

    // 合并HAnsi字段
    if original.HAnsi != "" && new.HAnsi != "" {
        merged.HAnsi = sm.mergeStringProperty(original.HAnsi, new.HAnsi, "fontHAnsi")
    } else if original.HAnsi != "" {
        merged.HAnsi = original.HAnsi
    } else if new.HAnsi != "" {
        merged.HAnsi = new.HAnsi
    }

    return merged
}

// mergeParagraphProperties 合并段落属性
func (sm *StyleMerger) mergeParagraphProperties(original, new *types.Paragraph) *types.Paragraph {
    if original == nil && new == nil {
        return nil
    }

    if original == nil {
        return new
    }

    if new == nil {
        return original
    }

    // 创建合并后的段落
    merged := &types.Paragraph{
        Text:       original.Text,
        Style:      original.Style,
        Runs:       original.Runs,
        HasComment: original.HasComment,
        CommentID:  original.CommentID,
    }

    // 合并文本
    if new.Text != "" {
        merged.Text = new.Text
    }

    // 合并样式
    if new.Style != "" {
        merged.Style = new.Style
    }

    // 合并运行
    if len(new.Runs) > 0 {
        merged.Runs = new.Runs
    }

    return merged
}

// mergeTableProperties 合并表格属性
func (sm *StyleMerger) mergeTableProperties(original, new *types.Table) *types.Table {
    if original == nil && new == nil {
        return nil
    }

    if original == nil {
        return new
    }

    if new == nil {
        return original
    }

    // 创建合并后的表格
    merged := &types.Table{
        Rows:    original.Rows,
        Columns: original.Columns,
    }

    // 合并行
    if len(new.Rows) > 0 {
        merged.Rows = new.Rows
    }

    // 合并列数
    if new.Columns > 0 {
        merged.Columns = new.Columns
    }

    return merged
}

func (sm *StyleMerger) getInheritanceChain(style *types.Style) []string {
    chain := []string{style.ID}
    current := style.BasedOn

    for current != "" {
        chain = append(chain, current)
        // 这里需要从样式管理器中获取基于的样式
        // 为了简化，我们假设基于的样式ID存在
        break // 避免无限循环
    }

    return chain
}

func (sm *StyleMerger) selectOptimalInheritanceChain(originalChain, newChain []string, priority int) []string {
    // 简单的选择策略：选择更长的继承链
    if len(newChain) > len(originalChain) {
        return newChain
    }
    return originalChain
}

func (sm *StyleMerger) mergeFormatProperties(original, new *types.StyleProperties, conflict *types.StyleConflict) *types.StyleProperties {
    merged := original.Clone()

    // 合并格式相关的属性
    if new.FontName != "" {
        merged.FontName = new.FontName
    }
    if new.FontSize > 0 {
        merged.FontSize = new.FontSize
    }
    if new.FontColor != "" {
        merged.FontColor = new.FontColor
    }
    if new.BackgroundColor != "" {
        merged.BackgroundColor = new.BackgroundColor
    }

    return merged
}

func (sm *StyleMerger) safeMergeProperties(target, source *types.StyleProperties) {
    // 安全地合并属性，避免覆盖现有值
    if source.FontName != "" && target.FontName == "" {
        target.FontName = source.FontName
    }
    if source.FontSize > 0 && target.FontSize == 0 {
        target.FontSize = source.FontSize
    }
    if source.FontColor != "" && target.FontColor == "" {
        target.FontColor = source.FontColor
    }
    if source.BackgroundColor != "" && target.BackgroundColor == "" {
        target.BackgroundColor = source.BackgroundColor
    }
}

// ValidateMergedStyle 验证合并后的样式
func (sm *StyleMerger) ValidateMergedStyle(style *types.Style) error {
    if style == nil {
        return fmt.Errorf("样式为空")
    }

    if style.ID == "" {
        return fmt.Errorf("样式ID为空")
    }

    if style.Name == "" {
        return fmt.Errorf("样式名称为空")
    }

    // 验证属性
    if style.Properties != nil {
        if err := sm.validateStyleProperties(style.Properties); err != nil {
            return fmt.Errorf("样式属性验证失败: %w", err)
        }
    }

    sm.logger.Info("样式验证通过，样式ID: %s, 名称: %s", style.ID, style.Name)

    return nil
}

func (sm *StyleMerger) validateStyleProperties(properties *types.StyleProperties) error {
    // 验证字体属性
    if properties.FontSize < 0 {
        return fmt.Errorf("字体大小不能为负数")
    }

    // 验证段落属性
    if properties.FirstLineIndent < 0 {
        return fmt.Errorf("首行缩进不能为负数")
    }
    if properties.LeftIndent < 0 {
        return fmt.Errorf("左缩进不能为负数")
    }
    if properties.RightIndent < 0 {
        return fmt.Errorf("右缩进不能为负数")
    }
    if properties.LineSpacing < 0 {
        return fmt.Errorf("行距不能为负数")
    }
    if properties.SpaceBefore < 0 {
        return fmt.Errorf("段前间距不能为负数")
    }
    if properties.SpaceAfter < 0 {
        return fmt.Errorf("段后间距不能为负数")
    }

    return nil
}

// mergeIntProperty 合并整数属性
func (sm *StyleMerger) mergeIntProperty(original, new int, propertyName string) int {
    if original == new {
        return original
    }

    // 记录冲突
    sm.logger.Warning("整数属性冲突，属性: %s, 原始: %d, 新: %d", propertyName, original, new)

    // 返回新值
    return new
}

// mergeBoolProperty 合并布尔属性
func (sm *StyleMerger) mergeBoolProperty(original, new bool, propertyName string) bool {
    if original == new {
        return original
    }

    // 记录冲突
    sm.logger.Warning("布尔属性冲突，属性: %s, 原始: %t, 新: %t", propertyName, original, new)

    // 返回新值
    return new
}
