package wordprocessingml

import (
	"fmt"
	"time"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/utils"
)

// StyleMerger 样式合并器
type StyleMerger struct {
	logger *utils.Logger
}

// NewStyleMerger 创建样式合并器
func NewStyleMerger() *StyleMerger {
	return &StyleMerger{
		logger: utils.NewLogger("StyleMerger"),
	}
}

// MergeStyles 合并样式
func (sm *StyleMerger) MergeStyles(conflict *types.StyleConflict) error {
	sm.logger.Info("开始合并样式", map[string]interface{}{
		"style_id":      conflict.StyleID,
		"conflict_type": conflict.Type,
		"priority":      conflict.Priority,
	})

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
	sm.logger.Info("合并属性冲突", map[string]interface{}{
		"style_id":        original.ID,
		"conflict_count":  len(conflict.ConflictingProperties),
	})

	// 创建合并后的样式
	mergedStyle := &types.Style{
		ID:         original.ID,
		Name:       original.Name,
		Type:       original.Type,
		BasedOn:    original.BasedOn,
		Next:       original.Next,
		Properties: &types.StyleProperties{},
		CreatedAt:  original.CreatedAt,
		UpdatedAt:  time.Now(),
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
	conflict.ResolvedAt = time.Now()
	conflict.Status = types.StyleConflictStatusResolved

	sm.logger.Info("属性冲突合并完成", map[string]interface{}{
		"style_id":           mergedStyle.ID,
		"merged_properties":  len(conflict.ConflictingProperties),
	})

	return nil
}

// mergeInheritanceConflicts 合并继承冲突
func (sm *StyleMerger) mergeInheritanceConflicts(original, new *types.Style, conflict *types.StyleConflict) error {
	sm.logger.Info("合并继承冲突", map[string]interface{}{
		"style_id":         original.ID,
		"original_based_on": original.BasedOn,
		"new_based_on":     new.BasedOn,
	})

	// 分析继承链
	originalChain := sm.getInheritanceChain(original)
	newChain := sm.getInheritanceChain(new)

	// 选择最优的继承链
	optimalChain := sm.selectOptimalInheritanceChain(originalChain, newChain, conflict.Priority)

	// 更新样式
	mergedStyle := original.Clone()
	mergedStyle.BasedOn = optimalChain[0]
	mergedStyle.UpdatedAt = time.Now()

	// 更新冲突结果
	conflict.ResolvedStyle = mergedStyle
	conflict.ResolvedAt = time.Now()
	conflict.Status = types.StyleConflictStatusResolved

	sm.logger.Info("继承冲突合并完成", map[string]interface{}{
		"style_id":      mergedStyle.ID,
		"new_based_on":  mergedStyle.BasedOn,
		"chain_length":  len(optimalChain),
	})

	return nil
}

// mergePriorityConflicts 合并优先级冲突
func (sm *StyleMerger) mergePriorityConflicts(original, new *types.Style, conflict *types.StyleConflict) error {
	sm.logger.Info("合并优先级冲突", map[string]interface{}{
		"style_id":         original.ID,
		"original_priority": conflict.OriginalPriority,
		"new_priority":     conflict.NewPriority,
	})

	// 根据优先级选择样式
	var selectedStyle *types.Style
	if conflict.NewPriority > conflict.OriginalPriority {
		selectedStyle = new
		sm.logger.Info("选择新样式（更高优先级）", map[string]interface{}{
			"style_id": new.ID,
			"priority": new.Priority,
		})
	} else {
		selectedStyle = original
		sm.logger.Info("保留原始样式（更高或相等优先级）", map[string]interface{}{
			"style_id": original.ID,
			"priority": original.Priority,
		})
	}

	// 更新冲突结果
	conflict.ResolvedStyle = selectedStyle
	conflict.ResolvedAt = time.Now()
	conflict.Status = types.StyleConflictStatusResolved

	return nil
}

// mergeFormatConflicts 合并格式冲突
func (sm *StyleMerger) mergeFormatConflicts(original, new *types.Style, conflict *types.StyleConflict) error {
	sm.logger.Info("合并格式冲突", map[string]interface{}{
		"style_id":       original.ID,
		"conflict_count": len(conflict.ConflictingProperties),
	})

	// 创建合并后的样式
	mergedStyle := original.Clone()
	mergedStyle.UpdatedAt = time.Now()

	// 合并格式属性
	if original.Properties != nil && new.Properties != nil {
		mergedStyle.Properties = sm.mergeFormatProperties(original.Properties, new.Properties, conflict)
	}

	// 更新冲突结果
	conflict.ResolvedStyle = mergedStyle
	conflict.ResolvedAt = time.Now()
	conflict.Status = types.StyleConflictStatusResolved

	sm.logger.Info("格式冲突合并完成", map[string]interface{}{
		"style_id":       mergedStyle.ID,
		"merged_formats": len(conflict.ConflictingProperties),
	})

	return nil
}

// mergeDefaultStrategy 默认合并策略
func (sm *StyleMerger) mergeDefaultStrategy(original, new *types.Style, conflict *types.StyleConflict) error {
	sm.logger.Info("使用默认合并策略", map[string]interface{}{
		"style_id": original.ID,
		"strategy": "conservative",
	})

	// 保守策略：保留原始样式，只添加新样式中不冲突的属性
	mergedStyle := original.Clone()
	mergedStyle.UpdatedAt = time.Now()

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
	conflict.ResolvedAt = time.Now()
	conflict.Status = types.StyleConflictStatusResolved

	return nil
}

// 辅助方法
func (sm *StyleMerger) mergeStringProperty(original, new, propertyName string) string {
	if new != "" && new != original {
		sm.logger.Debug("合并字符串属性", map[string]interface{}{
			"property": propertyName,
			"original": original,
			"new":      new,
		})
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

func (sm *StyleMerger) mergeStyleProperties(original, new *types.StyleProperties, conflict *types.StyleConflict) *types.StyleProperties {
	merged := &types.StyleProperties{}

	// 合并字体属性
	if original.Font != nil && new.Font != nil {
		merged.Font = sm.mergeFontProperties(original.Font, new.Font)
	} else if original.Font != nil {
		merged.Font = original.Font
	} else if new.Font != nil {
		merged.Font = new.Font
	}

	// 合并段落属性
	if original.Paragraph != nil && new.Paragraph != nil {
		merged.Paragraph = sm.mergeParagraphProperties(original.Paragraph, new.Paragraph)
	} else if original.Paragraph != nil {
		merged.Paragraph = original.Paragraph
	} else if new.Paragraph != nil {
		merged.Paragraph = new.Paragraph
	}

	// 合并表格属性
	if original.Table != nil && new.Table != nil {
		merged.Table = sm.mergeTableProperties(original.Table, new.Table)
	} else if original.Table != nil {
		merged.Table = original.Table
	} else if new.Table != nil {
		merged.Table = new.Table
	}

	return merged
}

func (sm *StyleMerger) mergeFontProperties(original, new *types.Font) *types.Font {
	merged := original.Clone()

	// 合并字体属性，新样式优先
	if new.Name != "" {
		merged.Name = new.Name
	}
	if new.Size > 0 {
		merged.Size = new.Size
	}
	if new.Bold {
		merged.Bold = new.Bold
	}
	if new.Italic {
		merged.Italic = new.Italic
	}
	if new.Underline {
		merged.Underline = new.Underline
	}
	if new.Color != "" {
		merged.Color = new.Color
	}

	return merged
}

func (sm *StyleMerger) mergeParagraphProperties(original, new *types.Paragraph) *types.Paragraph {
	merged := original.Clone()

	// 合并段落属性
	if new.Alignment != "" {
		merged.Alignment = new.Alignment
	}
	if new.Indent > 0 {
		merged.Indent = new.Indent
	}
	if new.Spacing > 0 {
		merged.Spacing = new.Spacing
	}

	return merged
}

func (sm *StyleMerger) mergeTableProperties(original, new *types.Table) *types.Table {
	merged := original.Clone()

	// 合并表格属性
	if new.Width > 0 {
		merged.Width = new.Width
	}
	if new.Height > 0 {
		merged.Height = new.Height
	}
	if new.Alignment != "" {
		merged.Alignment = new.Alignment
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
	if new.Font != nil {
		merged.Font = sm.mergeFontProperties(original.Font, new.Font)
	}
	if new.Paragraph != nil {
		merged.Paragraph = sm.mergeParagraphProperties(original.Paragraph, new.Paragraph)
	}

	return merged
}

func (sm *StyleMerger) safeMergeProperties(target, source *types.StyleProperties) {
	// 安全地合并属性，避免覆盖现有值
	if source.Font != nil && target.Font == nil {
		target.Font = source.Font.Clone()
	}
	if source.Paragraph != nil && target.Paragraph == nil {
		target.Paragraph = source.Paragraph.Clone()
	}
	if source.Table != nil && target.Table == nil {
		target.Table = source.Table.Clone()
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

	sm.logger.Info("样式验证通过", map[string]interface{}{
		"style_id": style.ID,
		"name":     style.Name,
	})

	return nil
}

func (sm *StyleMerger) validateStyleProperties(properties *types.StyleProperties) error {
	// 验证字体属性
	if properties.Font != nil {
		if properties.Font.Size < 0 {
			return fmt.Errorf("字体大小不能为负数")
		}
	}

	// 验证段落属性
	if properties.Paragraph != nil {
		if properties.Paragraph.Indent < 0 {
			return fmt.Errorf("段落缩进不能为负数")
		}
		if properties.Paragraph.Spacing < 0 {
			return fmt.Errorf("段落间距不能为负数")
		}
	}

	return nil
}
