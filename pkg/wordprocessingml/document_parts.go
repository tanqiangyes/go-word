// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// DocumentParts represents all parts of a Word document
type DocumentParts struct {
	// 主要文档部分
	MainDocumentPart *MainDocumentPart
	
	// 页眉页脚部分 (按使用频率排序)
	HeaderParts []HeaderPart
	FooterParts []FooterPart
	
	// 注释和脚注部分
	CommentParts  []CommentPart
	FootnoteParts []FootnotePart
	EndnoteParts  []EndnotePart
	
	// 样式和格式部分
	StylesPart    *StylesPart
	NumberingPart *NumberingPart
	SettingsPart  *SettingsPart
	
	// 主题和字体部分
	ThemePart     *ThemePart
	FontTablePart *FontTablePart
	
	// 其他部分
	WebSettingsPart *WebSettingsPart
	GlossaryPart    *GlossaryPart
	BibliographyPart *BibliographyPart
}

// MainDocumentPart represents the main document part
type MainDocumentPart struct {
	// 基础信息 (按大小排序)
	ID       string
	Content  *types.DocumentContent
	
	// 文档属性
	DocumentProperties map[string]interface{}
	
	// 关系
	Relationships []Relationship
}

// HeaderPart represents a header part
type HeaderPart struct {
	// 基础信息
	ID      string
	Type    HeaderFooterType
	Content []types.Paragraph
	
	// 属性
	Properties HeaderFooterProperties
	
	// 关系
	Relationships []Relationship
}

// FooterPart represents a footer part
type FooterPart struct {
	// 基础信息
	ID      string
	Type    HeaderFooterType
	Content []types.Paragraph
	
	// 属性
	Properties HeaderFooterProperties
	
	// 关系
	Relationships []Relationship
}

// CommentPart represents a comment part
type CommentPart struct {
	// 基础信息
	ID      string
	Content []Comment
	
	// 属性
	Properties CommentProperties
	
	// 关系
	Relationships []Relationship
}

// Comment represents a comment
type Comment struct {
	// 基础信息
	ID        string
	Author    string
	Date      string
	Text      string
	
	// 属性
	Initials  string
	Index     int
	ParentID  string
	
	// 格式
	Formatting CommentFormatting
}

// CommentProperties represents comment properties
type CommentProperties struct {
	// 基础设置
	Visible   bool
	Locked    bool
	Resolved  bool
	
	// 显示设置
	ShowAuthor bool
	ShowDate   bool
	ShowTime   bool
}

// CommentFormatting represents comment formatting
type CommentFormatting struct {
	// 字体属性
	FontName  string
	FontSize  int
	Bold      bool
	Italic    bool
	
	// 颜色
	Color     string
	Highlight string
}

// FootnotePart represents a footnote part
type FootnotePart struct {
	// 基础信息
	ID      string
	Content []Footnote
	
	// 属性
	Properties FootnoteProperties
	
	// 关系
	Relationships []Relationship
}

// Footnote represents a footnote
type Footnote struct {
	// 基础信息
	ID       string
	Text     string
	Number   int
	
	// 属性
	Type     FootnoteType
	Reference string
	
	// 格式
	Formatting FootnoteFormatting
}

// FootnoteType defines the type of footnote
type FootnoteType int

const (
	// NormalFootnote for normal footnotes
	NormalFootnote FootnoteType = iota
	// SeparatorFootnote for separator footnotes
	SeparatorFootnote
	// ContinuationSeparatorFootnote for continuation separator footnotes
	ContinuationSeparatorFootnote
	// ContinuationNoticeFootnote for continuation notice footnotes
	ContinuationNoticeFootnote
)

// FootnoteProperties represents footnote properties
type FootnoteProperties struct {
	// 基础设置
	RestartNumber    bool
	StartNumber      int
	NumberFormat     string
	
	// 位置设置
	Position         FootnotePosition
	Layout          FootnoteLayout
}

// FootnotePosition defines footnote position
type FootnotePosition int

const (
	// PageBottom for page bottom
	PageBottom FootnotePosition = iota
	// BeneathText for beneath text
	BeneathText
	// SectionBottom for section bottom
	SectionBottom
	// FootnoteDocumentEnd for document end
	FootnoteDocumentEnd
)

// FootnoteLayout represents footnote layout
type FootnoteLayout struct {
	// 布局设置
	Columns     int
	Spacing     float64
	Indent      float64
}

// FootnoteFormatting represents footnote formatting
type FootnoteFormatting struct {
	// 字体属性
	FontName  string
	FontSize  int
	Bold      bool
	Italic    bool
	
	// 颜色
	Color     string
}

// EndnotePart represents an endnote part
type EndnotePart struct {
	// 基础信息
	ID      string
	Content []Endnote
	
	// 属性
	Properties EndnoteProperties
	
	// 关系
	Relationships []Relationship
}

// Endnote represents an endnote
type Endnote struct {
	// 基础信息
	ID       string
	Text     string
	Number   int
	
	// 属性
	Type     EndnoteType
	Reference string
	
	// 格式
	Formatting EndnoteFormatting
}

// EndnoteType defines the type of endnote
type EndnoteType int

const (
	// NormalEndnote for normal endnotes
	NormalEndnote EndnoteType = iota
	// SeparatorEndnote for separator endnotes
	SeparatorEndnote
	// ContinuationSeparatorEndnote for continuation separator endnotes
	ContinuationSeparatorEndnote
	// ContinuationNoticeEndnote for continuation notice endnotes
	ContinuationNoticeEndnote
)

// EndnoteProperties represents endnote properties
type EndnoteProperties struct {
	// 基础设置
	RestartNumber    bool
	StartNumber      int
	NumberFormat     string
	
	// 位置设置
	Position         EndnotePosition
	Layout          EndnoteLayout
}

// EndnotePosition defines endnote position
type EndnotePosition int

const (
	// SectionEnd for section end
	SectionEnd EndnotePosition = iota
	// DocumentEnd for document end
	DocumentEnd
)

// EndnoteLayout represents endnote layout
type EndnoteLayout struct {
	// 布局设置
	Columns     int
	Spacing     float64
	Indent      float64
}

// EndnoteFormatting represents endnote formatting
type EndnoteFormatting struct {
	// 字体属性
	FontName  string
	FontSize  int
	Bold      bool
	Italic    bool
	
	// 颜色
	Color     string
}

// StylesPart represents styles part
type StylesPart struct {
	// 基础信息
	ID      string
	
	// 样式定义
	Styles  StyleDefinitions
	
	// 关系
	Relationships []Relationship
}

// StyleDefinitions represents style definitions
type StyleDefinitions struct {
	// 样式集合 (按使用频率排序)
	ParagraphStyles []ParagraphStyle
	CharacterStyles []CharacterStyle
	TableStyles     []TableStyle
	NumberingStyles []NumberingStyle
	
	// 默认样式
	DefaultStyles   DefaultStyleSet
	
	// 样式属性
	Properties      StyleProperties
}

// ParagraphStyle 使用 advanced_styles.go 中的定义

// ParagraphStyleProperties 使用 advanced_styles.go 中的定义

// CharacterStyle 使用 advanced_styles.go 中的定义

// CharacterStyleProperties 使用 advanced_styles.go 中的定义

// TableStyle 使用 advanced_styles.go 中的定义

// TableStyleProperties 使用 advanced_styles.go 中的定义

// NumberingStyle represents a numbering style
type NumberingStyle struct {
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
	Properties     NumberingStyleProperties
}

// NumberingStyleProperties 使用 advanced_styles.go 中的定义

// NumberingFormat represents numbering format
type NumberingFormat struct {
	// 基础设置
	Type        string
	Start       int
	Restart     bool
	
	// 格式设置
	Format      string
	Level       int
	Indent      float64
}

// DefaultStyleSet represents default styles
type DefaultStyleSet struct {
	// 默认样式
	Paragraph   string
	Character   string
	Table       string
	Numbering   string
	
	// 属性
	Properties  DefaultStyleProperties
}

// DefaultStyleProperties represents default style properties
type DefaultStyleProperties struct {
	// 基础设置
	Language    string
	Theme       string
	
	// 其他属性
	Hidden      bool
}

// StyleProperties represents style properties
type StyleProperties struct {
	// 基础设置
	Language    string
	Theme       string
	
	// 其他属性
	Hidden      bool
}

// NumberingPart represents numbering part
type NumberingPart struct {
	// 基础信息
	ID      string
	
	// 编号定义
	NumberingDefinitions []NumberingDefinition
	
	// 关系
	Relationships []Relationship
}

// NumberingDefinition represents a numbering definition
type NumberingDefinition struct {
	// 基础信息
	ID          string
	Name        string
	
	// 编号级别
	Levels      []NumberingLevel
	
	// 属性
	Properties  NumberingDefinitionProperties
}

// NumberingLevel represents a numbering level
type NumberingLevel struct {
	// 基础信息
	Index       int
	Start       int
	
	// 格式设置
	Format      string
	Alignment   string
	Indent      float64
	Hanging     float64
	
	// 属性
	Properties  NumberingLevelProperties
}

// NumberingLevelProperties represents numbering level properties
type NumberingLevelProperties struct {
	// 基础设置
	Restart     bool
	Legal       bool
	Legacy      bool
	
	// 其他属性
	Hidden      bool
}

// NumberingDefinitionProperties represents numbering definition properties
type NumberingDefinitionProperties struct {
	// 基础设置
	MultiLevel  bool
	Restart     bool
	
	// 其他属性
	Hidden      bool
}

// SettingsPart represents settings part
type SettingsPart struct {
	// 基础信息
	ID      string
	
	// 设置
	Settings DocumentSettings
	
	// 关系
	Relationships []Relationship
}

// DocumentSettings represents document settings
type DocumentSettings struct {
	// 视图设置
	ViewSettings     ViewSettings
	
	// 编辑设置
	EditSettings     EditSettings
	
	// 打印设置
	PrintSettings    PrintSettings
	
	// 其他设置
	OtherSettings    OtherSettings
}

// ViewSettings represents view settings
type ViewSettings struct {
	// 显示设置
	ShowWhiteSpace   bool
	ShowParagraphMarks bool
	ShowHiddenText   bool
	ShowBookmarks    bool
	
	// 缩放设置
	Zoom             int
	ZoomType         string
}

// EditSettings represents edit settings
type EditSettings struct {
	// 编辑限制
	TrackChanges     bool
	Protection       bool
	ReadOnly         bool
	
	// 其他设置
	AutoFormat       bool
	AutoCorrect      bool
}

// PrintSettings represents print settings
type PrintSettings struct {
	// 打印设置
	PrintBackground  bool
	PrintHiddenText  bool
	PrintComments    bool
	
	// 其他设置
	PrintProperties  bool
}

// OtherSettings represents other settings
type OtherSettings struct {
	// 其他设置
	UpdateFields     bool
	EmbedTrueTypeFonts bool
	SaveSubsetFonts  bool
	
	// 其他属性
	Hidden           bool
}

// ThemePart represents theme part
type ThemePart struct {
	// 基础信息
	ID      string
	
	// 主题
	Theme   Theme
	
	// 关系
	Relationships []Relationship
}

// Theme represents a theme
type Theme struct {
	// 基础信息
	Name        string
	Description string
	
	// 主题元素
	ColorScheme ColorScheme
	FontScheme  FontScheme
	FormatScheme FormatScheme
	EffectsScheme EffectsScheme
}

// ColorScheme represents a color scheme
type ColorScheme struct {
	// 基础信息
	Name        string
	
	// 颜色
	Colors      []ThemeColor
	AccentColors []ThemeColor
	HyperlinkColors []ThemeColor
}

// ThemeColor represents a theme color
type ThemeColor struct {
	// 基础信息
	Name        string
	Value       string
	
	// 属性
	Type        ColorType
	Index       int
}

// ColorType defines color type
type ColorType int

const (
	// PrimaryColor for primary colors
	PrimaryColor ColorType = iota
	// AccentColor for accent colors
	AccentColor
	// HyperlinkColor for hyperlink colors
	HyperlinkColor
)

// FontScheme represents a font scheme
type FontScheme struct {
	// 基础信息
	Name        string
	
	// 字体
	MajorFont   MajorFont
	MinorFont   MinorFont
}

// MajorFont represents major font
type MajorFont struct {
	// 基础信息
	Name        string
	Script      string
	
	// 字体
	LatinFont   string
	EastAsianFont string
	ComplexScriptFont string
}

// MinorFont represents minor font
type MinorFont struct {
	// 基础信息
	Name        string
	Script      string
	
	// 字体
	LatinFont   string
	EastAsianFont string
	ComplexScriptFont string
}

// FormatScheme represents a format scheme
type FormatScheme struct {
	// 基础信息
	Name        string
	
	// 格式
	Formats     []Format
}

// Format represents a format
type Format struct {
	// 基础信息
	Name        string
	Type        FormatType
	
	// 属性
	Properties  FormatProperties
}

// FormatType defines format type
type FormatType int

const (
	// FillFormat for fill formats
	FillFormat FormatType = iota
	// LineFormat for line formats
	LineFormat
	// EffectFormat for effect formats
	EffectFormat
)

// FormatProperties represents format properties
type FormatProperties struct {
	// 基础属性
	Visible     bool
	Locked      bool
	
	// 其他属性
	Hidden      bool
}

// EffectsScheme represents an effects scheme
type EffectsScheme struct {
	// 基础信息
	Name        string
	
	// 效果
	Effects     []Effect
}

// Effect represents an effect
type Effect struct {
	// 基础信息
	Name        string
	Type        EffectType
	
	// 属性
	Properties  EffectProperties
}

// EffectType 使用 advanced_text.go 中的定义
// EffectProperties 使用 advanced_text.go 中的定义

// FontTablePart represents font table part
type FontTablePart struct {
	// 基础信息
	ID      string
	
	// 字体表
	Fonts   []FontDefinition
	
	// 关系
	Relationships []Relationship
}

// FontDefinition represents a font definition
type FontDefinition struct {
	// 基础信息
	Name        string
	Family      string
	
	// 属性
	Pitch       FontPitch
	Charset     string
	Panose      string
	
	// 其他属性
	Hidden      bool
}

// FontPitch defines font pitch
type FontPitch int

const (
	// FixedPitch for fixed pitch fonts
	FixedPitch FontPitch = iota
	// VariablePitch for variable pitch fonts
	VariablePitch
	// DefaultPitch for default pitch fonts
	DefaultPitch
)

// WebSettingsPart represents web settings part
type WebSettingsPart struct {
	// 基础信息
	ID      string
	
	// 设置
	Settings WebSettings
	
	// 关系
	Relationships []Relationship
}

// WebSettings represents web settings
type WebSettings struct {
	// 基础设置
	Encoding       string
	TargetScreenSize string
	
	// 其他设置
	AllowPNG       bool
	RelyOnVML      bool
	AllowFonts     bool
}

// GlossaryPart represents glossary part
type GlossaryPart struct {
	// 基础信息
	ID      string
	
	// 内容
	Content []GlossaryEntry
	
	// 关系
	Relationships []Relationship
}

// GlossaryEntry represents a glossary entry
type GlossaryEntry struct {
	// 基础信息
	ID          string
	Term        string
	Definition  string
	
	// 属性
	Category    string
	Language    string
}

// BibliographyPart represents bibliography part
type BibliographyPart struct {
	// 基础信息
	ID      string
	
	// 内容
	Content []BibliographyEntry
	
	// 关系
	Relationships []Relationship
}

// BibliographyEntry represents a bibliography entry
type BibliographyEntry struct {
	// 基础信息
	ID          string
	Title       string
	Author      string
	Year        int
	
	// 属性
	Type        BibliographyType
	Source      string
}

// BibliographyType defines bibliography type
type BibliographyType int

const (
	// BookBibliography for books
	BookBibliography BibliographyType = iota
	// ArticleBibliography for articles
	ArticleBibliography
	// WebBibliography for web sources
	WebBibliography
)

// Relationship represents a relationship
type Relationship struct {
	// 基础信息
	ID          string
	Type        string
	Target      string
	
	// 属性
	TargetMode  string
}

// NewDocumentParts creates new document parts
func NewDocumentParts() *DocumentParts {
	return &DocumentParts{
		HeaderParts:    make([]HeaderPart, 0),
		FooterParts:    make([]FooterPart, 0),
		CommentParts:   make([]CommentPart, 0),
		FootnoteParts:  make([]FootnotePart, 0),
		EndnoteParts:   make([]EndnotePart, 0),
	}
}

// AddHeaderPart adds a header part
func (dp *DocumentParts) AddHeaderPart(header HeaderPart) {
	dp.HeaderParts = append(dp.HeaderParts, header)
}

// AddFooterPart adds a footer part
func (dp *DocumentParts) AddFooterPart(footer FooterPart) {
	dp.FooterParts = append(dp.FooterParts, footer)
}

// AddCommentPart adds a comment part
func (dp *DocumentParts) AddCommentPart(comment CommentPart) {
	dp.CommentParts = append(dp.CommentParts, comment)
}

// AddFootnotePart adds a footnote part
func (dp *DocumentParts) AddFootnotePart(footnote FootnotePart) {
	dp.FootnoteParts = append(dp.FootnoteParts, footnote)
}

// AddEndnotePart adds an endnote part
func (dp *DocumentParts) AddEndnotePart(endnote EndnotePart) {
	dp.EndnoteParts = append(dp.EndnoteParts, endnote)
}

// GetHeaderPart gets a header part by ID
func (dp *DocumentParts) GetHeaderPart(id string) *HeaderPart {
	for i := range dp.HeaderParts {
		if dp.HeaderParts[i].ID == id {
			return &dp.HeaderParts[i]
		}
	}
	return nil
}

// GetFooterPart gets a footer part by ID
func (dp *DocumentParts) GetFooterPart(id string) *FooterPart {
	for i := range dp.FooterParts {
		if dp.FooterParts[i].ID == id {
			return &dp.FooterParts[i]
		}
	}
	return nil
}

// GetCommentPart gets a comment part by ID
func (dp *DocumentParts) GetCommentPart(id string) *CommentPart {
	for i := range dp.CommentParts {
		if dp.CommentParts[i].ID == id {
			return &dp.CommentParts[i]
		}
	}
	return nil
}

// GetFootnotePart gets a footnote part by ID
func (dp *DocumentParts) GetFootnotePart(id string) *FootnotePart {
	for i := range dp.FootnoteParts {
		if dp.FootnoteParts[i].ID == id {
			return &dp.FootnoteParts[i]
		}
	}
	return nil
}

// GetEndnotePart gets an endnote part by ID
func (dp *DocumentParts) GetEndnotePart(id string) *EndnotePart {
	for i := range dp.EndnoteParts {
		if dp.EndnoteParts[i].ID == id {
			return &dp.EndnoteParts[i]
		}
	}
	return nil
}

// GetPartsSummary returns a summary of all parts
func (dp *DocumentParts) GetPartsSummary() string {
	var summary strings.Builder
	summary.WriteString("文档部分摘要:\n")
	summary.WriteString(fmt.Sprintf("页眉部分: %d\n", len(dp.HeaderParts)))
	summary.WriteString(fmt.Sprintf("页脚部分: %d\n", len(dp.FooterParts)))
	summary.WriteString(fmt.Sprintf("注释部分: %d\n", len(dp.CommentParts)))
	summary.WriteString(fmt.Sprintf("脚注部分: %d\n", len(dp.FootnoteParts)))
	summary.WriteString(fmt.Sprintf("尾注部分: %d\n", len(dp.EndnoteParts)))
	
	if dp.StylesPart != nil {
		summary.WriteString("样式部分: 已加载\n")
	}
	if dp.NumberingPart != nil {
		summary.WriteString("编号部分: 已加载\n")
	}
	if dp.SettingsPart != nil {
		summary.WriteString("设置部分: 已加载\n")
	}
	if dp.ThemePart != nil {
		summary.WriteString("主题部分: 已加载\n")
	}
	if dp.FontTablePart != nil {
		summary.WriteString("字体表部分: 已加载\n")
	}
	
	return summary.String()
}

// 添加缺失的类型定义

// CharacterStyle represents a character style
type CharacterStyle struct {
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
	Properties     CharacterStyleProperties
}

// CharacterStyleProperties 使用 advanced_styles.go 中的定义

// TableStyle represents a table style
type TableStyle struct {
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
	Properties     TableStyleProperties
}

// TableStyleProperties 使用 advanced_styles.go 中的定义

// Font 使用 format_support.go 中的定义
// TextEffects 使用 format_support.go 中的定义

// TableBorders 使用 advanced_formatting.go 中的定义
// BorderSide 使用 advanced_formatting.go 中的定义
// TableShading 使用 advanced_formatting.go 中的定义
// TableLayout 使用 advanced_formatting.go 中的定义
// CellProperties 使用 advanced_formatting.go 中的定义
// CellBorders 使用 advanced_formatting.go 中的定义
// CellShading 使用 advanced_formatting.go 中的定义
// CellMargin 使用 advanced_tables.go 中的定义

// HeaderFooterType 使用 advanced_formatting.go 中的定义
// HeaderFooterProperties 使用 advanced_formatting.go 中的定义
// RowProperties 使用 advanced_formatting.go 中的定义
// ColumnProperties 使用 advanced_formatting.go 中的定义

// CellProperties 使用 advanced_formatting.go 中的定义 