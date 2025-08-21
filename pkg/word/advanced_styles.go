// package word provides word document processing functionality
package word

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/utils"
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

	// 日志记录器
	Logger *utils.Logger
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
	Language string
	Theme    string
	Version  string

	// 其他属性
	Hidden bool
	Locked bool
}

// StyleDefinition represents a complete style definition
type StyleDefinition struct {
	// 基础信息
	ID       string
	Name     string
	Type     StyleType
	Category StyleCategory

	// 继承关系
	BasedOn string
	Next    string
	Link    string
	Parent  string

	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool
	Hidden         bool

	// 样式属性
	Properties *types.StyleProperties
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
	ID      string
	Name    string
	BasedOn string
	Next    string
	Link    string

	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool

	// 样式属性
	Properties *ParagraphStyleProperties
}

// ParagraphStyleProperties represents paragraph style properties
type ParagraphStyleProperties struct {
	// 段落属性
	Alignment string
	Indent    *IndentFormat
	Spacing   *SpacingFormat
	Borders   *BorderFormat
	Shading   *ShadingFormat
	Tabs      []TabStop

	// 文本属性
	Font    *Font
	Effects *TextEffects

	// 其他属性
	KeepLines       bool
	KeepNext        bool
	PageBreakBefore bool
	WidowControl    bool
	OutlineLevel    int
}

// CharacterStyleDefinition represents a character style definition
type CharacterStyleDefinition struct {
	// 基础信息
	ID      string
	Name    string
	BasedOn string
	Link    string

	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool

	// 样式属性
	Properties *CharacterStyleProperties
}

// CharacterStyleProperties represents character style properties
type CharacterStyleProperties struct {
	// 字体属性
	FontName        string
	FontSize        int
	FontColor       string
	BackgroundColor string
	Bold            bool
	Italic          bool
	Underline       bool
	StrikeThrough   bool

	// 其他属性
	Hidden     bool
	Vanish     bool
	SpecVanish bool
	Display    string
}

// TableStyleDefinition represents a table style definition
type TableStyleDefinition struct {
	// 基础信息
	ID      string
	Name    string
	BasedOn string
	Next    string

	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool

	// 样式属性
	Properties *TableStyleProperties
}

// TableStyleProperties represents table style properties
type TableStyleProperties struct {
	// 表格属性
	Borders   *TableBorders
	Shading   *TableShading
	Layout    *TableLayout
	Alignment string

	// 单元格属性
	CellProperties *CellProperties

	// 其他属性
	Hidden       bool
	AllowOverlap bool
	AllowBreak   bool
}

// NumberingStyleDefinition represents a numbering style definition
type NumberingStyleDefinition struct {
	// 基础信息
	ID      string
	Name    string
	BasedOn string

	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool

	// 样式属性
	Properties *NumberingStyleProperties
}

// NumberingStyleProperties represents numbering style properties
type NumberingStyleProperties struct {
	// 编号属性
	Numbering *NumberingFormat
	Indent    *IndentFormat
	Alignment string

	// 其他属性
	Hidden  bool
	Restart bool
	Legal   bool
}

// ListStyleDefinition represents a list style definition
type ListStyleDefinition struct {
	// 基础信息
	ID      string
	Name    string
	BasedOn string

	// 属性
	SemiHidden     bool
	UnhideWhenUsed bool
	QFormat        bool
	Locked         bool

	// 样式属性
	Properties *ListStyleProperties
}

// ListStyleProperties represents list style properties
type ListStyleProperties struct {
	// 列表属性
	ListType ListType
	Levels   []ListLevel
	Restart  bool

	// 其他属性
	Hidden bool
	Legal  bool
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
	Index int
	Start int

	// 格式设置
	Format     string
	Alignment  string
	Indent     float64
	Hanging    float64
	TextIndent float64

	// 属性
	Restart bool
	Legal   bool
	Legacy  bool
}

// TabStop represents a tab stop
type TabStop struct {
	// 基础信息
	Position  float64
	Alignment TabAlignment
	Leader    TabLeader
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
	Conflicts []types.StyleConflict

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
	ID       string
	Name     string
	Priority int

	// 规则条件
	Condition string
	Action    string

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
		StyleCache:       make(map[string]*StyleDefinition),
		InheritanceChain: make(map[string][]string),
		ConflictResolver: &StyleConflictResolver{
			ResolutionStrategy: KeepOriginalStrategy,
			Conflicts:          make([]types.StyleConflict, 0),
			Rules:              make([]ConflictResolutionRule, 0),
		},
		Logger: &utils.Logger{},
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
		ID:             style.ID,
		Name:           style.Name,
		Type:           ParagraphStyleType,
		Category:       CustomCategory,
		BasedOn:        style.BasedOn,
		Next:           style.Next,
		Link:           style.Link,
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
		ID:             style.ID,
		Name:           style.Name,
		Type:           CharacterStyleType,
		Category:       CustomCategory,
		BasedOn:        style.BasedOn,
		Link:           style.Link,
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
		ID:             style.ID,
		Name:           style.Name,
		Type:           TableStyleType,
		Category:       CustomCategory,
		BasedOn:        style.BasedOn,
		Next:           style.Next,
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

// convertTypesStyleToStyleDefinition converts types.Style to StyleDefinition
func (ass *AdvancedStyleSystem) convertTypesStyleToStyleDefinition(style *types.Style) *StyleDefinition {
	if style == nil {
		return nil
	}

	return &StyleDefinition{
		ID:             style.ID,
		Name:           style.Name,
		Type:           ass.convertStringToStyleType(string(style.Type)),
		Category:       CustomCategory,
		BasedOn:        "",
		Next:           "",
		Link:           "",
		Parent:         "",
		SemiHidden:     false,
		UnhideWhenUsed: false,
		QFormat:        false,
		Locked:         false,
		Hidden:         false,
		Properties:     &types.StyleProperties{},
	}
}

// convertStringToStyleType converts string to StyleType
func (ass *AdvancedStyleSystem) convertStringToStyleType(styleType string) StyleType {
	switch styleType {
	case "paragraph":
		return ParagraphStyleType
	case "character":
		return CharacterStyleType
	case "table":
		return TableStyleType
	case "numbering":
		return NumberingStyleType
	case "list":
		return ListStyleType
	default:
		return ParagraphStyleType
	}
}

// convertStyleTypeToString converts StyleType to string
func (ass *AdvancedStyleSystem) convertStyleTypeToString(styleType StyleType) string {
	switch styleType {
	case ParagraphStyleType:
		return "paragraph"
	case CharacterStyleType:
		return "character"
	case TableStyleType:
		return "table"
	case NumberingStyleType:
		return "numbering"
	case ListStyleType:
		return "list"
	default:
		return "paragraph"
	}
}

// checkStyleConflict checks for style conflicts
func (ass *AdvancedStyleSystem) checkStyleConflict(id string, styleType StyleType) *types.StyleConflict {
	existingStyle := ass.StyleCache[id]
	if existingStyle == nil {
		return nil
	}

	// 检查同名冲突
	return &types.StyleConflict{
		ID:          fmt.Sprintf("conflict_%s", id),
		StyleName:   id,
		Type:        types.StyleConflictTypeProperty,
		Description: fmt.Sprintf("Style conflict detected for %s", id),
		Severity:    "medium",
		Resolved:    false,
		Priority:    0,
		OriginalStyle: &types.Style{
			ID:   existingStyle.ID,
			Name: existingStyle.Name,
			Type: types.StyleType(ass.convertStyleTypeToString(existingStyle.Type)),
		},
	}
}

// resolveStyleConflict resolves a style conflict
func (ass *AdvancedStyleSystem) resolveStyleConflict(conflict *types.StyleConflict) error {
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
func (ass *AdvancedStyleSystem) mergeStyles(conflict *types.StyleConflict) error {
	ass.Logger.Info("开始合并样式，样式ID: %s, 冲突类型: %s, 优先级: %d", conflict.StyleID, conflict.Type, conflict.Priority)

	// 获取冲突的样式
	originalStyle := ass.GetStyle(conflict.StyleID)
	if originalStyle == nil {
		return fmt.Errorf("原始样式未找到: %s", conflict.StyleID)
	}

	newStyle := ass.convertTypesStyleToStyleDefinition(conflict.NewStyle)
	if newStyle == nil {
		return fmt.Errorf("新样式为空")
	}

	// 根据冲突类型选择合并策略
	switch conflict.Type {
	case types.StyleConflictTypeProperty:
		return ass.mergePropertyConflicts(originalStyle, newStyle, conflict)
	case types.StyleConflictTypeInheritance:
		return ass.mergeInheritanceConflicts(originalStyle, newStyle, conflict)
	case types.StyleConflictTypePriority:
		return ass.mergePriorityConflicts(originalStyle, newStyle, conflict)
	case types.StyleConflictTypeFormat:
		return ass.mergeFormatConflicts(originalStyle, newStyle, conflict)
	default:
		return ass.mergeDefaultStrategy(originalStyle, newStyle, conflict)
	}
}

// mergePropertyConflicts 合并属性冲突
func (ass *AdvancedStyleSystem) mergePropertyConflicts(original, new *StyleDefinition, conflict *types.StyleConflict) error {
	ass.Logger.Info("合并属性冲突，样式ID: %s, 冲突数量: %d", original.ID, len(conflict.ConflictingProperties))

	// 创建合并后的样式
	mergedStyle := &StyleDefinition{
		ID:             original.ID,
		Name:           original.Name,
		Type:           original.Type,
		BasedOn:        original.BasedOn,
		Next:           original.Next,
		Link:           original.Link,
		Parent:         original.Parent,
		SemiHidden:     original.SemiHidden,
		UnhideWhenUsed: original.UnhideWhenUsed,
		QFormat:        original.QFormat,
		Locked:         original.Locked,
		Hidden:         original.Hidden,
		Properties:     &types.StyleProperties{},
	}

	// 合并样式属性
	if original.Properties != nil && new.Properties != nil {
		mergedStyle.Properties = ass.mergeStyleProperties(original.Properties, new.Properties, conflict)
	} else if original.Properties != nil {
		mergedStyle.Properties = original.Properties
	} else if new.Properties != nil {
		mergedStyle.Properties = new.Properties
	}

	// 更新样式缓存
	ass.StyleCache[mergedStyle.ID] = mergedStyle

	// 记录合并结果
	ass.Logger.Info("属性冲突合并完成，样式ID: %s, 合并属性数: %d", mergedStyle.ID, len(conflict.ConflictingProperties))

	return nil
}

// mergeInheritanceConflicts 合并继承冲突
func (ass *AdvancedStyleSystem) mergeInheritanceConflicts(original, new *StyleDefinition, conflict *types.StyleConflict) error {
	ass.Logger.Info("合并继承冲突，样式ID: %s, 原始继承: %s, 新继承: %s", original.ID, original.BasedOn, new.BasedOn)

	// 分析继承链
	originalChain := ass.getInheritanceChain(original.ID)
	newChain := ass.getInheritanceChain(new.ID)

	// 选择最优的继承链
	optimalChain := ass.selectOptimalInheritanceChain(originalChain, newChain, conflict.Priority)

	// 更新样式
	mergedStyle := &StyleDefinition{
		ID:             original.ID,
		Name:           original.Name,
		Type:           original.Type,
		BasedOn:        optimalChain[0],
		Next:           original.Next,
		Link:           original.Link,
		Parent:         original.Parent,
		SemiHidden:     original.SemiHidden,
		UnhideWhenUsed: original.UnhideWhenUsed,
		QFormat:        original.QFormat,
		Locked:         original.Locked,
		Hidden:         original.Hidden,
		Properties:     original.Properties,
	}

	// 更新继承链
	ass.InheritanceChain[mergedStyle.ID] = optimalChain

	// 更新样式缓存
	ass.StyleCache[mergedStyle.ID] = mergedStyle

	ass.Logger.Info("继承冲突合并完成，样式ID: %s, 新继承: %s, 链长度: %d", mergedStyle.ID, mergedStyle.BasedOn, len(optimalChain))

	return nil
}

// mergePriorityConflicts 合并优先级冲突
func (ass *AdvancedStyleSystem) mergePriorityConflicts(original, new *StyleDefinition, conflict *types.StyleConflict) error {
	ass.Logger.Info("合并优先级冲突，样式ID: %s, 原始优先级: %d, 新优先级: %d", original.ID, conflict.OriginalPriority, conflict.NewPriority)

	// 根据优先级选择样式
	var selectedStyle *StyleDefinition
	if conflict.NewPriority > conflict.OriginalPriority {
		selectedStyle = new
		ass.Logger.Info("选择新样式（更高优先级），样式ID: %s", new.ID)
	} else {
		selectedStyle = original
		ass.Logger.Info("保留原始样式（更高或相等优先级），样式ID: %s", original.ID)
	}

	// 更新样式缓存
	ass.StyleCache[selectedStyle.ID] = selectedStyle

	return nil
}

// mergeFormatConflicts 合并格式冲突
func (ass *AdvancedStyleSystem) mergeFormatConflicts(original, new *StyleDefinition, conflict *types.StyleConflict) error {
	ass.Logger.Info("合并格式冲突，样式ID: %s, 冲突数量: %d", original.ID, len(conflict.ConflictingProperties))

	// 创建合并后的样式
	mergedStyle := &StyleDefinition{
		ID:             original.ID,
		Name:           original.Name,
		Type:           original.Type,
		BasedOn:        original.BasedOn,
		Next:           original.Next,
		Link:           original.Link,
		Parent:         original.Parent,
		SemiHidden:     original.SemiHidden,
		UnhideWhenUsed: original.UnhideWhenUsed,
		QFormat:        original.QFormat,
		Locked:         original.Locked,
		Hidden:         original.Hidden,
		Properties:     original.Properties,
	}

	// 合并格式属性
	if original.Properties != nil && new.Properties != nil {
		mergedStyle.Properties = ass.mergeFormatProperties(original.Properties, new.Properties, conflict)
	}

	// 更新样式缓存
	ass.StyleCache[mergedStyle.ID] = mergedStyle

	ass.Logger.Info("格式冲突合并完成，样式ID: %s, 合并格式数: %d", mergedStyle.ID, len(conflict.ConflictingProperties))

	return nil
}

// mergeDefaultStrategy 默认合并策略
func (ass *AdvancedStyleSystem) mergeDefaultStrategy(original, new *StyleDefinition, conflict *types.StyleConflict) error {
	ass.Logger.Info("使用默认合并策略，样式ID: %s, 策略: %s", original.ID, "conservative")

	// 保守策略：保留原始样式，只添加新样式中不冲突的属性
	mergedStyle := &StyleDefinition{
		ID:             original.ID,
		Name:           original.Name,
		Type:           original.Type,
		BasedOn:        original.BasedOn,
		Next:           original.Next,
		Link:           original.Link,
		Parent:         original.Parent,
		SemiHidden:     original.SemiHidden,
		UnhideWhenUsed: original.UnhideWhenUsed,
		QFormat:        original.QFormat,
		Locked:         original.Locked,
		Hidden:         original.Hidden,
		Properties:     original.Properties,
	}

	// 合并非冲突属性
	if new.Properties != nil {
		if mergedStyle.Properties == nil {
			mergedStyle.Properties = &types.StyleProperties{}
		}

		// 安全地合并属性
		ass.safeMergeProperties(mergedStyle.Properties, new.Properties)
	}

	// 更新样式缓存
	ass.StyleCache[mergedStyle.ID] = mergedStyle

	return nil
}

// 辅助方法
func (ass *AdvancedStyleSystem) mergeStringProperty(original, new, propertyName string) string {
	if new != "" && new != original {
		ass.Logger.Debug("合并字符串属性，属性: %s, 原始: %s, 新: %s", propertyName, original, new)
		return new
	}
	return original
}

func (ass *AdvancedStyleSystem) mergeStringSlice(original, new []string) []string {
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

// mergeIntProperty merges integer properties
func (ass *AdvancedStyleSystem) mergeIntProperty(original, new int, propertyName string) int {
	if new > 0 && new != original {
		ass.Logger.Debug("合并整数属性，属性: %s, 原始: %d, 新: %d", propertyName, original, new)
		return new
	}
	return original
}

// mergeFloatProperty merges float properties
func (ass *AdvancedStyleSystem) mergeFloatProperty(original, new float64, propertyName string) float64 {
	if new > 0 && new != original {
		ass.Logger.Debug("合并浮点数属性，属性: %s, 原始: %f, 新: %f", propertyName, original, new)
		return new
	}
	return original
}

// mergeBoolProperty merges boolean properties
func (ass *AdvancedStyleSystem) mergeBoolProperty(original, new bool, propertyName string) bool {
	if new != original {
		ass.Logger.Debug("合并布尔属性，属性: %s, 原始: %t, 新: %t", propertyName, original, new)
		return new
	}
	return original
}

func (ass *AdvancedStyleSystem) mergeStyleProperties(original, new *types.StyleProperties, conflict *types.StyleConflict) *types.StyleProperties {
	merged := &types.StyleProperties{}

	// 合并字体属性
	merged.FontName = ass.mergeStringProperty(original.FontName, new.FontName, "fontName")
	merged.FontSize = ass.mergeIntProperty(original.FontSize, new.FontSize, "fontSize")
	merged.FontColor = ass.mergeStringProperty(original.FontColor, new.FontColor, "fontColor")
	merged.BackgroundColor = ass.mergeStringProperty(original.BackgroundColor, new.BackgroundColor, "backgroundColor")
	merged.Bold = ass.mergeBoolProperty(original.Bold, new.Bold, "bold")
	merged.Italic = ass.mergeBoolProperty(original.Italic, new.Italic, "italic")
	merged.Underline = ass.mergeBoolProperty(original.Underline, new.Underline, "underline")
	merged.StrikeThrough = ass.mergeBoolProperty(original.StrikeThrough, new.StrikeThrough, "strikeThrough")

	// 合并段落属性
	merged.Alignment = ass.mergeStringProperty(original.Alignment, new.Alignment, "alignment")
	merged.LineSpacing = ass.mergeFloatProperty(original.LineSpacing, new.LineSpacing, "lineSpacing")
	merged.SpaceBefore = ass.mergeFloatProperty(original.SpaceBefore, new.SpaceBefore, "spaceBefore")
	merged.SpaceAfter = ass.mergeFloatProperty(original.SpaceAfter, new.SpaceAfter, "spaceAfter")
	merged.FirstLineIndent = ass.mergeFloatProperty(original.FirstLineIndent, new.FirstLineIndent, "firstLineIndent")
	merged.LeftIndent = ass.mergeFloatProperty(original.LeftIndent, new.LeftIndent, "leftIndent")
	merged.RightIndent = ass.mergeFloatProperty(original.RightIndent, new.RightIndent, "rightIndent")
	merged.KeepLines = ass.mergeBoolProperty(original.KeepLines, new.KeepLines, "keepLines")
	merged.KeepNext = ass.mergeBoolProperty(original.KeepNext, new.KeepNext, "keepNext")
	merged.PageBreakBefore = ass.mergeBoolProperty(original.PageBreakBefore, new.PageBreakBefore, "pageBreakBefore")
	merged.WidowControl = ass.mergeBoolProperty(original.WidowControl, new.WidowControl, "widowControl")

	// 表格属性在 types.StyleProperties 中不可用，跳过

	return merged
}

// 此函数已被重写为使用 types.StyleProperties，不再需要

// 此函数已被重写为使用 types.StyleProperties，不再需要

// 此函数已被重写为使用 types.StyleProperties，不再需要

func (ass *AdvancedStyleSystem) getInheritanceChain(styleID string) []string {
	if chain, exists := ass.InheritanceChain[styleID]; exists {
		return chain
	}
	return []string{styleID}
}

func (ass *AdvancedStyleSystem) selectOptimalInheritanceChain(originalChain, newChain []string, priority int) []string {
	// 简单的选择策略：选择更长的继承链
	if len(newChain) > len(originalChain) {
		return newChain
	}
	return originalChain
}

func (ass *AdvancedStyleSystem) mergeFormatProperties(original, new *types.StyleProperties, conflict *types.StyleConflict) *types.StyleProperties {
	merged := &types.StyleProperties{}

	// 合并格式相关的属性
	merged.FontName = ass.mergeStringProperty(original.FontName, new.FontName, "fontName")
	merged.FontSize = ass.mergeIntProperty(original.FontSize, new.FontSize, "fontSize")
	merged.FontColor = ass.mergeStringProperty(original.FontColor, new.FontColor, "fontColor")
	merged.BackgroundColor = ass.mergeStringProperty(original.BackgroundColor, new.BackgroundColor, "backgroundColor")
	merged.Bold = ass.mergeBoolProperty(original.Bold, new.Bold, "bold")
	merged.Italic = ass.mergeBoolProperty(original.Italic, new.Italic, "italic")
	merged.Underline = ass.mergeBoolProperty(original.Underline, new.Underline, "underline")
	merged.StrikeThrough = ass.mergeBoolProperty(original.StrikeThrough, new.StrikeThrough, "strikeThrough")

	// 合并段落属性
	merged.Alignment = ass.mergeStringProperty(original.Alignment, new.Alignment, "alignment")
	merged.LineSpacing = ass.mergeFloatProperty(original.LineSpacing, new.LineSpacing, "lineSpacing")
	merged.SpaceBefore = ass.mergeFloatProperty(original.SpaceBefore, new.SpaceBefore, "spaceBefore")
	merged.SpaceAfter = ass.mergeFloatProperty(original.SpaceAfter, new.SpaceAfter, "spaceAfter")
	merged.FirstLineIndent = ass.mergeFloatProperty(original.FirstLineIndent, new.FirstLineIndent, "firstLineIndent")
	merged.LeftIndent = ass.mergeFloatProperty(original.LeftIndent, new.LeftIndent, "leftIndent")
	merged.RightIndent = ass.mergeFloatProperty(original.RightIndent, new.RightIndent, "rightIndent")
	merged.KeepLines = ass.mergeBoolProperty(original.KeepLines, new.KeepLines, "keepLines")
	merged.KeepNext = ass.mergeBoolProperty(original.KeepNext, new.KeepNext, "keepNext")
	merged.PageBreakBefore = ass.mergeBoolProperty(original.PageBreakBefore, new.PageBreakBefore, "pageBreakBefore")
	merged.WidowControl = ass.mergeBoolProperty(original.WidowControl, new.WidowControl, "widowControl")

	return merged
}

func (ass *AdvancedStyleSystem) safeMergeProperties(target, source *types.StyleProperties) {
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
	if source.Bold && !target.Bold {
		target.Bold = source.Bold
	}
	if source.Italic && !target.Italic {
		target.Italic = source.Italic
	}
	if source.Underline && !target.Underline {
		target.Underline = source.Underline
	}
	if source.StrikeThrough && !target.StrikeThrough {
		target.StrikeThrough = source.StrikeThrough
	}
	if source.Alignment != "" && target.Alignment == "" {
		target.Alignment = source.Alignment
	}
	if source.LineSpacing > 0 && target.LineSpacing == 0 {
		target.LineSpacing = source.LineSpacing
	}
	if source.SpaceBefore > 0 && target.SpaceBefore == 0 {
		target.SpaceBefore = source.SpaceBefore
	}
	if source.SpaceAfter > 0 && target.SpaceAfter == 0 {
		target.SpaceAfter = source.SpaceAfter
	}
	if source.FirstLineIndent > 0 && target.FirstLineIndent == 0 {
		target.FirstLineIndent = source.FirstLineIndent
	}
	if source.LeftIndent > 0 && target.LeftIndent == 0 {
		target.LeftIndent = source.LeftIndent
	}
	if source.RightIndent > 0 && target.RightIndent == 0 {
		target.RightIndent = source.RightIndent
	}
	if source.KeepLines && !target.KeepLines {
		target.KeepLines = source.KeepLines
	}
	if source.KeepNext && !target.KeepNext {
		target.KeepNext = source.KeepNext
	}
	if source.PageBreakBefore && !target.PageBreakBefore {
		target.PageBreakBefore = source.PageBreakBefore
	}
	if source.WidowControl && !target.WidowControl {
		target.WidowControl = source.WidowControl
	}
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
		if style.Properties.FontName != "" {
			run.FontName = style.Properties.FontName
		}
		if style.Properties.FontSize > 0 {
			run.FontSize = style.Properties.FontSize
		}
		if style.Properties.Bold {
			run.Bold = true
		}
		if style.Properties.Italic {
			run.Italic = true
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
