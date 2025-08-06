// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// AdvancedStyleSystem represents the advanced style system
type AdvancedStyleSystem struct {
	// 样式管理器
	StyleManager *StyleManager
	
	// 样式缓存
	StyleCache map[string]*StyleDefinition
	
	// 样式继承链
	InheritanceChain map[string][]string
	
	// 样式冲突解决器
	ConflictResolver *StyleConflictResolver
}

// StyleManager manages all styles in the document
type StyleManager struct {
	// 样式集合 (按使用频率排序)
	ParagraphStyles map[string]*ParagraphStyleDefinition
	CharacterStyles map[string]*CharacterStyleDefinition
	TableStyles     map[string]*TableStyleDefinition
	NumberingStyles map[string]*NumberingStyleDefinition
	ListStyles      map[string]*ListStyleDefinition
	
	// 默认样式
	DefaultStyles *DefaultStyleSet
	
	// 样式属性
	Properties *StyleManagerProperties
}

// StyleManagerProperties represents style manager properties
type StyleManagerProperties struct {
	// 基础设置
	Language    string
	Theme       string
	Version     string
	
	// 其他属性
	Hidden      bool
	Locked      bool
}

// StyleDefinition represents a complete style definition
type StyleDefinition struct {
	// 基础信息
	ID          string
	Name        string
	Type        StyleType
	Category    StyleCategory
	
	// 继承关系
	BasedOn     string
	Next        string
	Link        string
	Parent      string
	
	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool
	Hidden         bool
	
	// 样式属性
	Properties     *StyleProperties
}

// StyleType defines the type of style
type StyleType int

const (
	// ParagraphStyleType for paragraph styles
	ParagraphStyleType StyleType = iota
	// CharacterStyleType for character styles
	CharacterStyleType
	// TableStyleType for table styles
	TableStyleType
	// NumberingStyleType for numbering styles
	NumberingStyleType
	// ListStyleType for list styles
	ListStyleType
)

// StyleCategory defines the category of style
type StyleCategory int

const (
	// BuiltInCategory for built-in styles
	BuiltInCategory StyleCategory = iota
	// CustomCategory for custom styles
	CustomCategory
	// UserCategory for user-defined styles
	UserCategory
)

// ParagraphStyleDefinition represents a paragraph style definition
type ParagraphStyleDefinition struct {
	// 基础信息
	ID          string
	Name        string
	BasedOn     string
	Next        string
	Link        string
	
	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool
	
	// 样式属性
	Properties     *ParagraphStyleProperties
}

// ParagraphStyleProperties represents paragraph style properties
type ParagraphStyleProperties struct {
	// 段落属性
	Alignment      string
	Indent         *IndentFormat
	Spacing        *SpacingFormat
	Borders        *BorderFormat
	Shading        *ShadingFormat
	Tabs           []TabStop
	
	// 文本属性
	Font           *Font
	Effects        *TextEffects
	
	// 其他属性
	KeepLines      bool
	KeepNext       bool
	PageBreakBefore bool
	WidowControl   bool
	OutlineLevel   int
}

// CharacterStyleDefinition represents a character style definition
type CharacterStyleDefinition struct {
	// 基础信息
	ID          string
	Name        string
	BasedOn     string
	Link        string
	
	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool
	
	// 样式属性
	Properties     *CharacterStyleProperties
}

// CharacterStyleProperties represents character style properties
type CharacterStyleProperties struct {
	// 字体属性
	Font           *Font
	Effects        *TextEffects
	
	// 其他属性
	Hidden         bool
	Vanish         bool
	SpecVanish     bool
	Display        string
}

// TableStyleDefinition represents a table style definition
type TableStyleDefinition struct {
	// 基础信息
	ID          string
	Name        string
	BasedOn     string
	Next        string
	
	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool
	
	// 样式属性
	Properties     *TableStyleProperties
}

// TableStyleProperties represents table style properties
type TableStyleProperties struct {
	// 表格属性
	Borders        *TableBorders
	Shading        *TableShading
	Layout          *TableLayout
	Alignment      string
	
	// 单元格属性
	CellProperties  *CellProperties
	
	// 其他属性
	Hidden         bool
	AllowOverlap   bool
	AllowBreak     bool
}

// NumberingStyleDefinition represents a numbering style definition
type NumberingStyleDefinition struct {
	// 基础信息
	ID          string
	Name        string
	BasedOn     string
	
	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool
	
	// 样式属性
	Properties     *NumberingStyleProperties
}

// NumberingStyleProperties represents numbering style properties
type NumberingStyleProperties struct {
	// 编号属性
	Numbering      *NumberingFormat
	Indent         *IndentFormat
	Alignment      string
	
	// 其他属性
	Hidden         bool
	Restart        bool
	Legal          bool
}

// ListStyleDefinition represents a list style definition
type ListStyleDefinition struct {
	// 基础信息
	ID          string
	Name        string
	BasedOn     string
	
	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool
	
	// 样式属性
	Properties     *ListStyleProperties
}

// ListStyleProperties represents list style properties
type ListStyleProperties struct {
	// 列表属性
	ListType       ListType
	Levels         []ListLevel
	Restart        bool
	
	// 其他属性
	Hidden         bool
	Legal          bool
}

// ListType defines the type of list
type ListType int

const (
	// BulletList for bullet lists
	BulletList ListType = iota
	// NumberedList for numbered lists
	NumberedList
	// MixedList for mixed lists
	MixedList
)

// ListLevel represents a list level
type ListLevel struct {
	// 基础信息
	Index       int
	Start       int
	
	// 格式设置
	Format      string
	Alignment   string
	Indent      float64
	Hanging     float64
	TextIndent  float64
	
	// 属性
	Restart     bool
	Legal       bool
	Legacy      bool
}

// TabStop represents a tab stop
type TabStop struct {
	// 基础信息
	Position    float64
	Alignment   TabAlignment
	Leader      TabLeader
}

// TabAlignment defines tab alignment
type TabAlignment int

const (
	// LeftTab for left alignment
	LeftTab TabAlignment = iota
	// CenterTab for center alignment
	CenterTab
	// RightTab for right alignment
	RightTab
	// DecimalTab for decimal alignment
	DecimalTab
	// BarTab for bar alignment
	BarTab
)

// TabLeader defines tab leader
type TabLeader int

const (
	// NoLeader for no leader
	NoLeader TabLeader = iota
	// DotLeader for dot leader
	DotLeader
	// DashLeader for dash leader
	DashLeader
	// UnderlineLeader for underline leader
	UnderlineLeader
	// HeavyLeader for heavy leader
	HeavyLeader
	// MiddleDotLeader for middle dot leader
	MiddleDotLeader
)

// StyleConflictResolver resolves style conflicts
type StyleConflictResolver struct {
	// 冲突解决策略
	ResolutionStrategy ConflictResolutionStrategy
	
	// 冲突记录
	Conflicts []StyleConflict
	
	// 解决规则
	Rules []ConflictResolutionRule
}

// ConflictResolutionStrategy defines conflict resolution strategy
type ConflictResolutionStrategy int

const (
	// KeepOriginalStrategy keeps original style
	KeepOriginalStrategy ConflictResolutionStrategy = iota
	// UseNewerStrategy uses newer style
	UseNewerStrategy
	// MergeStrategy merges styles
	MergeStrategy
	// UserChoiceStrategy lets user choose
	UserChoiceStrategy
)

// StyleConflict represents a style conflict
type StyleConflict struct {
	// 基础信息
	ID          string
	StyleName   string
	ConflictType ConflictType
	
	// 冲突详情
	OriginalStyle *StyleDefinition
	NewStyle      *StyleDefinition
	Resolution    *StyleDefinition
	
	// 属性
	Resolved      bool
	ResolutionDate string
}

// ConflictType defines the type of conflict
type ConflictType int

const (
	// PropertyConflict for property conflicts
	PropertyConflict ConflictType = iota
	// InheritanceConflict for inheritance conflicts
	InheritanceConflict
	// NamingConflict for naming conflicts
	NamingConflict
	// FormattingConflict for formatting conflicts
	FormattingConflict
)

// ConflictResolutionRule represents a conflict resolution rule
type ConflictResolutionRule struct {
	// 基础信息
	ID          string
	Name        string
	Priority    int
	
	// 规则条件
	Condition   string
	Action      string
	
	// 属性
	Enabled     bool
	Description string
}

// NewAdvancedStyleSystem creates a new advanced style system
func NewAdvancedStyleSystem() *AdvancedStyleSystem {
	return &AdvancedStyleSystem{
		StyleManager: &StyleManager{
			ParagraphStyles: make(map[string]*ParagraphStyleDefinition),
			CharacterStyles: make(map[string]*CharacterStyleDefinition),
			TableStyles:     make(map[string]*TableStyleDefinition),
			NumberingStyles: make(map[string]*NumberingStyleDefinition),
			ListStyles:      make(map[string]*ListStyleDefinition),
			DefaultStyles:   &DefaultStyleSet{},
			Properties:      &StyleManagerProperties{},
		},
		StyleCache:        make(map[string]*StyleDefinition),
		InheritanceChain:  make(map[string][]string),
		ConflictResolver:  &StyleConflictResolver{
			ResolutionStrategy: KeepOriginalStrategy,
			Conflicts:          make([]StyleConflict, 0),
			Rules:              make([]ConflictResolutionRule, 0),
		},
	}
}

// AddParagraphStyle adds a paragraph style
func (ass *AdvancedStyleSystem) AddParagraphStyle(style *ParagraphStyleDefinition) error {
	if style == nil {
		return fmt.Errorf("style cannot be nil")
	}
	
	if style.ID == "" {
		return fmt.Errorf("style ID cannot be empty")
	}
	
	// 检查冲突
	if conflict := ass.checkStyleConflict(style.ID, ParagraphStyleType); conflict != nil {
		return ass.resolveStyleConflict(conflict)
	}
	
	// 添加到样式管理器
	ass.StyleManager.ParagraphStyles[style.ID] = style
	
	// 添加到缓存
	ass.StyleCache[style.ID] = &StyleDefinition{
		ID:          style.ID,
		Name:        style.Name,
		Type:        ParagraphStyleType,
		Category:    CustomCategory,
		BasedOn:     style.BasedOn,
		Next:        style.Next,
		Link:        style.Link,
		SemiHidden:     style.SemiHidden,
		UnhideWhenUsed: style.UnhideWhenUsed,
		QFormat:        style.QFormat,
		Locked:         style.Locked,
	}
	
	// 更新继承链
	ass.updateInheritanceChain(style.ID, style.BasedOn)
	
	return nil
}

// AddCharacterStyle adds a character style
func (ass *AdvancedStyleSystem) AddCharacterStyle(style *CharacterStyleDefinition) error {
	if style == nil {
		return fmt.Errorf("style cannot be nil")
	}
	
	if style.ID == "" {
		return fmt.Errorf("style ID cannot be empty")
	}
	
	// 检查冲突
	if conflict := ass.checkStyleConflict(style.ID, CharacterStyleType); conflict != nil {
		return ass.resolveStyleConflict(conflict)
	}
	
	// 添加到样式管理器
	ass.StyleManager.CharacterStyles[style.ID] = style
	
	// 添加到缓存
	ass.StyleCache[style.ID] = &StyleDefinition{
		ID:          style.ID,
		Name:        style.Name,
		Type:        CharacterStyleType,
		Category:    CustomCategory,
		BasedOn:     style.BasedOn,
		Link:        style.Link,
		SemiHidden:     style.SemiHidden,
		UnhideWhenUsed: style.UnhideWhenUsed,
		QFormat:        style.QFormat,
		Locked:         style.Locked,
	}
	
	// 更新继承链
	ass.updateInheritanceChain(style.ID, style.BasedOn)
	
	return nil
}

// AddTableStyle adds a table style
func (ass *AdvancedStyleSystem) AddTableStyle(style *TableStyleDefinition) error {
	if style == nil {
		return fmt.Errorf("style cannot be nil")
	}
	
	if style.ID == "" {
		return fmt.Errorf("style ID cannot be empty")
	}
	
	// 检查冲突
	if conflict := ass.checkStyleConflict(style.ID, TableStyleType); conflict != nil {
		return ass.resolveStyleConflict(conflict)
	}
	
	// 添加到样式管理器
	ass.StyleManager.TableStyles[style.ID] = style
	
	// 添加到缓存
	ass.StyleCache[style.ID] = &StyleDefinition{
		ID:          style.ID,
		Name:        style.Name,
		Type:        TableStyleType,
		Category:    CustomCategory,
		BasedOn:     style.BasedOn,
		Next:        style.Next,
		SemiHidden:     style.SemiHidden,
		UnhideWhenUsed: style.UnhideWhenUsed,
		QFormat:        style.QFormat,
		Locked:         style.Locked,
	}
	
	// 更新继承链
	ass.updateInheritanceChain(style.ID, style.BasedOn)
	
	return nil
}

// GetStyle gets a style by ID
func (ass *AdvancedStyleSystem) GetStyle(id string) *StyleDefinition {
	return ass.StyleCache[id]
}

// GetParagraphStyle gets a paragraph style by ID
func (ass *AdvancedStyleSystem) GetParagraphStyle(id string) *ParagraphStyleDefinition {
	return ass.StyleManager.ParagraphStyles[id]
}

// GetCharacterStyle gets a character style by ID
func (ass *AdvancedStyleSystem) GetCharacterStyle(id string) *CharacterStyleDefinition {
	return ass.StyleManager.CharacterStyles[id]
}

// GetTableStyle gets a table style by ID
func (ass *AdvancedStyleSystem) GetTableStyle(id string) *TableStyleDefinition {
	return ass.StyleManager.TableStyles[id]
}

// GetInheritanceChain gets the inheritance chain for a style
func (ass *AdvancedStyleSystem) GetInheritanceChain(id string) []string {
	return ass.InheritanceChain[id]
}

// checkStyleConflict checks for style conflicts
func (ass *AdvancedStyleSystem) checkStyleConflict(id string, styleType StyleType) *StyleConflict {
	existingStyle := ass.StyleCache[id]
	if existingStyle == nil {
		return nil
	}
	
	// 检查同名冲突
	return &StyleConflict{
		ID:          fmt.Sprintf("conflict_%s", id),
		StyleName:   id,
		ConflictType: NamingConflict,
		OriginalStyle: existingStyle,
		Resolved:    false,
	}
}

// resolveStyleConflict resolves a style conflict
func (ass *AdvancedStyleSystem) resolveStyleConflict(conflict *StyleConflict) error {
	switch ass.ConflictResolver.ResolutionStrategy {
	case KeepOriginalStrategy:
		return fmt.Errorf("style conflict: keeping original style %s", conflict.StyleName)
	case UseNewerStrategy:
		// 使用新样式，移除旧样式
		delete(ass.StyleCache, conflict.StyleName)
		return nil
	case MergeStrategy:
		// 合并样式
		return ass.mergeStyles(conflict)
	case UserChoiceStrategy:
		// 添加到冲突列表，等待用户选择
		ass.ConflictResolver.Conflicts = append(ass.ConflictResolver.Conflicts, *conflict)
		return fmt.Errorf("style conflict: waiting for user choice for style %s", conflict.StyleName)
	default:
		return fmt.Errorf("unknown conflict resolution strategy")
	}
}

// mergeStyles merges conflicting styles
func (ass *AdvancedStyleSystem) mergeStyles(conflict *StyleConflict) error {
	// 简单的合并策略：保留原始样式，添加新样式的非冲突属性
	// 这里可以实现更复杂的合并逻辑
	return fmt.Errorf("style merging not implemented yet")
}

// updateInheritanceChain updates the inheritance chain for a style
func (ass *AdvancedStyleSystem) updateInheritanceChain(id, basedOn string) {
	if basedOn == "" {
		return
	}
	
	chain := []string{id}
	current := basedOn
	
	for current != "" {
		chain = append(chain, current)
		if style := ass.StyleCache[current]; style != nil {
			current = style.BasedOn
		} else {
			break
		}
	}
	
	ass.InheritanceChain[id] = chain
}

// GetStyleSummary returns a summary of all styles
func (ass *AdvancedStyleSystem) GetStyleSummary() string {
	var summary strings.Builder
	summary.WriteString("高级样式系统摘要:\n")
	summary.WriteString(fmt.Sprintf("段落样式: %d\n", len(ass.StyleManager.ParagraphStyles)))
	summary.WriteString(fmt.Sprintf("字符样式: %d\n", len(ass.StyleManager.CharacterStyles)))
	summary.WriteString(fmt.Sprintf("表格样式: %d\n", len(ass.StyleManager.TableStyles)))
	summary.WriteString(fmt.Sprintf("编号样式: %d\n", len(ass.StyleManager.NumberingStyles)))
	summary.WriteString(fmt.Sprintf("列表样式: %d\n", len(ass.StyleManager.ListStyles)))
	summary.WriteString(fmt.Sprintf("样式缓存: %d\n", len(ass.StyleCache)))
	summary.WriteString(fmt.Sprintf("继承链: %d\n", len(ass.InheritanceChain)))
	summary.WriteString(fmt.Sprintf("冲突数量: %d\n", len(ass.ConflictResolver.Conflicts)))
	
	return summary.String()
}

// ApplyStyle applies a style to content
func (ass *AdvancedStyleSystem) ApplyStyle(content interface{}, styleID string) error {
	style := ass.GetStyle(styleID)
	if style == nil {
		return fmt.Errorf("style not found: %s", styleID)
	}
	
	switch content := content.(type) {
	case *types.Paragraph:
		return ass.applyParagraphStyle(content, styleID)
	case *types.Run:
		return ass.applyCharacterStyle(content, styleID)
	case *types.Table:
		return ass.applyTableStyle(content, styleID)
	default:
		return fmt.Errorf("unsupported content type for style application")
	}
}

// applyParagraphStyle applies a paragraph style
func (ass *AdvancedStyleSystem) applyParagraphStyle(paragraph *types.Paragraph, styleID string) error {
	style := ass.GetParagraphStyle(styleID)
	if style == nil {
		return fmt.Errorf("paragraph style not found: %s", styleID)
	}
	
	paragraph.Style = styleID
	
	// 应用样式属性
	if style.Properties != nil {
		if style.Properties.Alignment != "" {
			// 这里可以设置段落对齐方式
		}
		if style.Properties.KeepLines {
			// 这里可以设置保持行在一起
		}
		if style.Properties.KeepNext {
			// 这里可以设置与下一段保持在一起
		}
		if style.Properties.PageBreakBefore {
			// 这里可以设置段前分页
		}
		if style.Properties.WidowControl {
			// 这里可以设置孤行控制
		}
	}
	
	return nil
}

// applyCharacterStyle applies a character style
func (ass *AdvancedStyleSystem) applyCharacterStyle(run *types.Run, styleID string) error {
	style := ass.GetCharacterStyle(styleID)
	if style == nil {
		return fmt.Errorf("character style not found: %s", styleID)
	}
	
	// 应用样式属性
	if style.Properties != nil {
		if style.Properties.Font != nil {
			if style.Properties.Font.Name != "" {
				run.FontName = style.Properties.Font.Name
			}
			if style.Properties.Font.Size > 0 {
				run.FontSize = int(style.Properties.Font.Size)
			}
			if style.Properties.Font.Bold {
				run.Bold = true
			}
			if style.Properties.Font.Italic {
				run.Italic = true
			}
		}
	}
	
	return nil
}

// applyTableStyle applies a table style
func (ass *AdvancedStyleSystem) applyTableStyle(table *types.Table, styleID string) error {
	style := ass.GetTableStyle(styleID)
	if style == nil {
		return fmt.Errorf("table style not found: %s", styleID)
	}
	
	// 应用样式属性
	if style.Properties != nil {
		if style.Properties.Alignment != "" {
			// 这里可以设置表格对齐方式
		}
		if style.Properties.Hidden {
			// 这里可以设置表格隐藏
		}
	}
	
	return nil
} 