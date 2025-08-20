// Package word provides word document processing functionality
package word

import (
	"fmt"
	"os"
	"time"

	"github.com/tanqiangyes/go-word/pkg/opc"
	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/utils"
)

// DocumentBuilder provides a fluent interface for building documents
type DocumentBuilder struct {
	Document *Document
	Config   *DocumentConfig
	Logger   *utils.Logger
}

// DocumentConfig holds configuration options for document creation
type DocumentConfig struct {
	Title      string
	Author     string
	Subject    string
	Keywords   []string
	Category   string
	Comments   string
	Language   string
	Template   string
	Protection *types.DocumentProtectionConfig
	Formatting *FormattingConfig
	Validation *types.DocumentValidationConfig
}

// ProtectionConfig holds document protection settings
type ProtectionConfig struct {
	Enabled        bool
	Password       string
	ProtectionType ProtectionType
	Permissions    map[string]bool
	Watermark      *WatermarkConfig
}

// WatermarkConfig holds watermark settings
type WatermarkConfig struct {
	Text         string
	Font         string
	Size         int
	Color        string
	Transparency float64
	Rotation     float64
}

// FormattingConfig holds document formatting settings
type FormattingConfig struct {
	DefaultFont     string
	DefaultFontSize int
	LineSpacing     float64
	Margins         *MarginConfig
	PageSize        *PageSizeConfig
	Theme           string
}

// MarginConfig holds page margin settings
type MarginConfig struct {
	Top    float64
	Bottom float64
	Left   float64
	Right  float64
}

// PageSizeConfig holds page size settings
type PageSizeConfig struct {
	Width       float64
	Height      float64
	Orientation string
}

// ValidationConfig holds document validation settings
type ValidationConfig struct {
	Enabled    bool
	Rules      []ValidationRule
	AutoFix    bool
	StrictMode bool
}

// NewDocumentBuilder creates a new document builder with fluent interface
func NewDocumentBuilder() *DocumentBuilder {
	return &DocumentBuilder{
		Config: &DocumentConfig{
			Language: "zh-CN",
			Formatting: &FormattingConfig{
				DefaultFont:     "Microsoft YaHei",
				DefaultFontSize: 12,
				LineSpacing:     1.15,
				Margins: &MarginConfig{
					Top: 72, Bottom: 72, Left: 72, Right: 72,
				},
				PageSize: &PageSizeConfig{
					Width: 595, Height: 842, Orientation: "portrait",
				},
			},
			Protection: &types.DocumentProtectionConfig{
				Type:    types.ProtectionTypeNone,
				Enabled: false,
			},
			Validation: &types.DocumentValidationConfig{
				ValidateStructure: true,
				ValidateContent:   true,
				ValidateStyles:    true,
				Enabled:           true,
				AutoFix:           false,
				StrictMode:        false,
			},
		},
		Logger: utils.NewLogger(utils.LogLevelInfo, os.Stdout),
	}
}

// WithTitle sets the document title
func (b *DocumentBuilder) WithTitle(title string) *DocumentBuilder {
	b.Config.Title = title
	return b
}

// WithAuthor sets the document author
func (b *DocumentBuilder) WithAuthor(author string) *DocumentBuilder {
	b.Config.Author = author
	return b
}

// WithSubject sets the document subject
func (b *DocumentBuilder) WithSubject(subject string) *DocumentBuilder {
	b.Config.Subject = subject
	return b
}

// WithKeywords sets the document keywords
func (b *DocumentBuilder) WithKeywords(keywords ...string) *DocumentBuilder {
	b.Config.Keywords = keywords
	return b
}

// WithLanguage sets the document language
func (b *DocumentBuilder) WithLanguage(language string) *DocumentBuilder {
	b.Config.Language = language
	return b
}

// WithTemplate sets the document template
func (b *DocumentBuilder) WithTemplate(template string) *DocumentBuilder {
	b.Config.Template = template
	return b
}

// WithProtection enables document protection
func (b *DocumentBuilder) WithProtection(protectionType types.ProtectionType, password string) *DocumentBuilder {
	b.Config.Protection.Enabled = true
	b.Config.Protection.Type = protectionType
	b.Config.Protection.Password = password
	return b
}

// WithPermissions sets document permissions
func (b *DocumentBuilder) WithPermissions(permissions map[string]bool) *DocumentBuilder {
	// Convert map[string]bool to []string
	permList := make([]string, 0, len(permissions))
	for perm, enabled := range permissions {
		if enabled {
			permList = append(permList, perm)
		}
	}
	b.Config.Protection.Permissions = permList
	return b
}

// WithWatermark adds a watermark to the document
func (b *DocumentBuilder) WithWatermark(text, font string, size int, color string) *DocumentBuilder {
	b.Config.Protection.Watermark = &types.WatermarkConfig{
		Text:         text,
		Font:         font,
		Size:         size,
		Color:        color,
		Transparency: 0.5,
		Rotation:     45,
	}
	return b
}

// WithDefaultFont sets the default font
func (b *DocumentBuilder) WithDefaultFont(font string, size int) *DocumentBuilder {
	b.Config.Formatting.DefaultFont = font
	b.Config.Formatting.DefaultFontSize = size
	return b
}

// WithMargins sets the page margins
func (b *DocumentBuilder) WithMargins(top, bottom, left, right float64) *DocumentBuilder {
	b.Config.Formatting.Margins.Top = top
	b.Config.Formatting.Margins.Bottom = bottom
	b.Config.Formatting.Margins.Left = left
	b.Config.Formatting.Margins.Right = right
	return b
}

// WithPageSize sets the page size
func (b *DocumentBuilder) WithPageSize(width, height float64, orientation string) *DocumentBuilder {
	b.Config.Formatting.PageSize.Width = width
	b.Config.Formatting.PageSize.Height = height
	b.Config.Formatting.PageSize.Orientation = orientation
	return b
}

// WithTheme sets the document theme
func (b *DocumentBuilder) WithTheme(theme string) *DocumentBuilder {
	b.Config.Formatting.Theme = theme
	return b
}

// WithValidation enables document validation
func (b *DocumentBuilder) WithValidation(enabled, autoFix, strictMode bool) *DocumentBuilder {
	b.Config.Validation.Enabled = enabled
	b.Config.Validation.AutoFix = autoFix
	b.Config.Validation.StrictMode = strictMode
	return b
}

// Build creates the document with the current configuration
func (b *DocumentBuilder) Build() (*Document, error) {
	// Create new document
	doc := &Document{}

	// Initialize document parts
	doc.documentParts = NewDocumentParts()

	// Apply configuration
	if err := b.applyConfiguration(doc); err != nil {
		return nil, fmt.Errorf("failed to apply configuration: %w", err)
	}

	b.Document = doc
	return doc, nil
}

// applyConfiguration applies the builder configuration to the document
func (b *DocumentBuilder) applyConfiguration(doc *Document) error {
	// Set document properties
	if b.Config.Title != "" {
		if err := b.setDocumentTitle(doc, b.Config.Title); err != nil {
			return fmt.Errorf("设置文档标题失败: %w", err)
		}
	}

	if b.Config.Author != "" {
		if err := b.setDocumentAuthor(doc, b.Config.Author); err != nil {
			return fmt.Errorf("设置文档作者失败: %w", err)
		}
	}

	// Apply protection if enabled
	if b.Config.Protection.Type != types.ProtectionTypeNone {
		if err := b.applyDocumentProtection(doc, *b.Config.Protection); err != nil {
			return fmt.Errorf("应用文档保护失败: %w", err)
		}
	}

	// Apply validation if enabled
	if b.Config.Validation.ValidateStructure || b.Config.Validation.ValidateContent || b.Config.Validation.ValidateStyles {
		if err := b.applyDocumentValidation(doc, *b.Config.Validation); err != nil {
			return fmt.Errorf("应用文档验证失败: %w", err)
		}
	}

	return nil
}

// setDocumentTitle 设置文档标题
func (b *DocumentBuilder) setDocumentTitle(doc *Document, title string) error {
	// 设置文档核心属性中的标题
	if doc.coreProperties == nil {
		doc.coreProperties = &types.CoreProperties{}
	}

	doc.coreProperties.Title = title

	// 同时更新文档元数据
	if doc.metadata == nil {
		doc.metadata = make(map[string]interface{})
	}
	doc.metadata["title"] = title

	b.Logger.Info("文档标题已设置: %s", title)

	return nil
}

// setDocumentAuthor 设置文档作者
func (b *DocumentBuilder) setDocumentAuthor(doc *Document, author string) error {
	// 设置文档核心属性中的作者
	if doc.coreProperties == nil {
		doc.coreProperties = &types.CoreProperties{}
	}

	doc.coreProperties.Creator = author
	doc.coreProperties.LastModifiedBy = author

	// 同时更新文档元数据
	if doc.metadata == nil {
		doc.metadata = make(map[string]interface{})
	}
	doc.metadata["author"] = author
	doc.metadata["creator"] = author

	b.Logger.Info("文档作者已设置: %s", author)

	return nil
}

// applyDocumentProtection 应用文档保护
func (b *DocumentBuilder) applyDocumentProtection(doc *Document, protection types.DocumentProtectionConfig) error {
	// 简化的文档保护实现
	if protection.Enabled && protection.Type != types.ProtectionTypeNone {
		// 设置文档保护标志
		if doc.metadata == nil {
			doc.metadata = make(map[string]interface{})
		}
		doc.metadata["protection"] = map[string]interface{}{
			"type":     protection.Type,
			"password": protection.Password != "",
			"enabled":  protection.Enabled,
		}

		b.Logger.Info("文档保护已应用，保护类型: %s, 启用: %t", protection.Type, protection.Enabled)
	}

	return nil
}

// applyDocumentValidation 应用文档验证
func (b *DocumentBuilder) applyDocumentValidation(doc *Document, validation types.DocumentValidationConfig) error {
	// 简化的文档验证实现
	if validation.Enabled {
		// 设置文档验证标志
		if doc.metadata == nil {
			doc.metadata = make(map[string]interface{})
		}
		doc.metadata["validation"] = map[string]interface{}{
			"validateStructure": validation.ValidateStructure,
			"validateContent":   validation.ValidateContent,
			"validateStyles":    validation.ValidateStyles,
			"enabled":           validation.Enabled,
			"autoFix":           validation.AutoFix,
			"strictMode":        validation.StrictMode,
		}

		b.Logger.Info("文档验证已应用，验证结构: %t, 验证内容: %t, 验证样式: %t", validation.ValidateStructure, validation.ValidateContent, validation.ValidateStyles)
	}

	return nil
}

// ParagraphBuilder provides a fluent interface for building paragraphs
type ParagraphBuilder struct {
	Paragraph types.Paragraph
	Content   []types.Run
	Logger    *utils.Logger
}

// NewParagraphBuilder creates a new paragraph builder
func NewParagraphBuilder() *ParagraphBuilder {
	return &ParagraphBuilder{
		Paragraph: types.Paragraph{
			Runs: make([]types.Run, 0),
		},
		Logger: utils.NewLogger(utils.LogLevelInfo, os.Stdout),
	}
}

// WithText adds plain text to the paragraph
func (b *ParagraphBuilder) WithText(text string) *ParagraphBuilder {
	b.Paragraph.Text = text
	b.Paragraph.Runs = append(b.Paragraph.Runs, types.Run{
		Text: text,
	})
	return b
}

// WithFormattedText adds formatted text to the paragraph
func (b *ParagraphBuilder) WithFormattedText(text string, formatting *TextFormatting) *ParagraphBuilder {
	run := types.Run{
		Text: text,
	}

	if formatting != nil {
		run.Bold = formatting.Bold
		run.Italic = formatting.Italic
		run.Underline = formatting.Underline
		run.FontSize = formatting.FontSize
		run.FontName = formatting.FontName
		run.Color = formatting.Color
	}

	b.Paragraph.Runs = append(b.Paragraph.Runs, run)
	return b
}

// WithStyle sets the paragraph style
func (b *ParagraphBuilder) WithStyle(style string) *ParagraphBuilder {
	b.Paragraph.Style = style
	return b
}

// WithComment adds a comment to the paragraph
func (b *ParagraphBuilder) WithComment(author, text string) *ParagraphBuilder {
	b.Paragraph.HasComment = true
	b.Paragraph.CommentID = fmt.Sprintf("comment_%d", time.Now().Unix())

	// 简化的评论处理
	b.Logger.Info("评论已添加，评论ID: %s, 作者: %s, 文本: %s", b.Paragraph.CommentID, author, text)

	return b
}

// Build creates the paragraph
func (b *ParagraphBuilder) Build() types.Paragraph {
	return b.Paragraph
}

// TextFormatting holds text formatting options
type TextFormatting struct {
	Bold      bool
	Italic    bool
	Underline bool
	FontSize  int
	FontName  string
	Color     string
}

// TableBuilder provides a fluent interface for building tables
type TableBuilder struct {
	Table   types.Table
	Rows    [][]string
	Headers []string
	Style   string
}

// NewTableBuilder creates a new table builder
func NewTableBuilder() *TableBuilder {
	return &TableBuilder{
		Table: types.Table{
			Rows: make([]types.TableRow, 0),
		},
	}
}

// WithHeaders sets the table headers
func (b *TableBuilder) WithHeaders(headers ...string) *TableBuilder {
	b.Headers = headers
	b.Table.Columns = len(headers)
	return b
}

// WithRows adds rows to the table
func (b *TableBuilder) WithRows(rows ...[]string) *TableBuilder {
	b.Rows = append(b.Rows, rows...)
	return b
}

// WithStyle sets the table style
func (b *TableBuilder) WithStyle(style string) *TableBuilder {
	b.Style = style
	return b
}

// Build creates the table
func (b *TableBuilder) Build() types.Table {
	// Add header row if headers are provided
	if len(b.Headers) > 0 {
		headerRow := types.TableRow{
			Cells: make([]types.TableCell, len(b.Headers)),
		}
		for i, header := range b.Headers {
			headerRow.Cells[i] = types.TableCell{
				Text: header,
			}
		}
		b.Table.Rows = append(b.Table.Rows, headerRow)
	}

	// Add data rows
	for _, rowData := range b.Rows {
		row := types.TableRow{
			Cells: make([]types.TableCell, len(rowData)),
		}
		for i, cellData := range rowData {
			row.Cells[i] = types.TableCell{
				Text: cellData,
			}
		}
		b.Table.Rows = append(b.Table.Rows, row)
	}

	return b.Table
}

// DocumentOperations provides a fluent interface for document operations
type DocumentOperations struct {
	Document *Document
}

// NewDocumentOperations creates document operations for a document
func NewDocumentOperations(doc *Document) *DocumentOperations {
	return &DocumentOperations{
		Document: doc,
	}
}

// AddParagraph adds a paragraph to the document
func (ops *DocumentOperations) AddParagraph(builder *ParagraphBuilder) *DocumentOperations {
	paragraph := builder.Build()
	
	// 确保文档主部分已初始化
	if ops.Document.mainPart == nil {
		ops.Document.mainPart = &MainDocumentPart{
			Content: &types.DocumentContent{
				Paragraphs: []types.Paragraph{},
				Tables:     []types.Table{},
			},
		}
	}
	
	// 确保文档内容已初始化
	if ops.Document.mainPart.Content == nil {
		ops.Document.mainPart.Content = &types.DocumentContent{
			Paragraphs: make([]types.Paragraph, 0),
			Tables:     make([]types.Table, 0),
		}
	}
	
	// 添加段落到文档
	ops.Document.mainPart.Content.Paragraphs = append(ops.Document.mainPart.Content.Paragraphs, paragraph)
	
	return ops
}

// AddTable adds a table to the document
func (ops *DocumentOperations) AddTable(builder *TableBuilder) *DocumentOperations {
	table := builder.Build()
	
	// 确保文档主部分已初始化
	if ops.Document.mainPart == nil {
		ops.Document.mainPart = &MainDocumentPart{
			Content: &types.DocumentContent{
				Paragraphs: []types.Paragraph{},
				Tables:     []types.Table{},
			},
		}
	}
	
	// 确保文档内容已初始化
	if ops.Document.mainPart.Content == nil {
		ops.Document.mainPart.Content = &types.DocumentContent{
			Paragraphs: make([]types.Paragraph, 0),
			Tables:     make([]types.Table, 0),
		}
	}
	
	// 添加表格到文档
	ops.Document.mainPart.Content.Tables = append(ops.Document.mainPart.Content.Tables, table)
	
	return ops
}

// Save saves the document to a file
func (ops *DocumentOperations) Save(filename string) error {
	// 创建OPC容器
	container, err := opc.New()
	if err != nil {
		return fmt.Errorf("创建OPC容器失败: %w", err)
	}
	defer container.Close()

	// 构建文档内容
	content := ops.buildDocumentContent()

	// 添加到容器
	container.AddPart("word/document.xml", []byte(content), "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")

	// 保存文件
	return container.SaveToFile(filename)
}

// buildDocumentContent builds the document XML content
func (ops *DocumentOperations) buildDocumentContent() string {
	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>`
	
	// 检查文档主部分是否存在
	if ops.Document.mainPart != nil && ops.Document.mainPart.Content != nil {
		// 添加段落
		for _, paragraph := range ops.Document.mainPart.Content.Paragraphs {
			content += "<w:p><w:r><w:t>" + paragraph.Text + "</w:t></w:r></w:p>"
		}
		
		// 添加表格
		for _, table := range ops.Document.mainPart.Content.Tables {
			content += "<w:tbl>"
			for _, row := range table.Rows {
				content += "<w:tr>"
				for _, cell := range row.Cells {
					content += "<w:tc><w:p><w:r><w:t>" + cell.Text + "</w:t></w:r></w:p></w:tc>"
				}
				content += "</w:tr>"
			}
			content += "</w:tbl>"
		}
	}
	
	content += "</w:body></w:document>"
	return content
}

// FluentDocument provides a fluent interface for document operations
type FluentDocument struct {
	Document   *Document
	Operations *DocumentOperations
}

// NewFluentDocument creates a new fluent document interface
func NewFluentDocument(doc *Document) *FluentDocument {
	return &FluentDocument{
		Document:   doc,
		Operations: NewDocumentOperations(doc),
	}
}

// AddParagraph adds a paragraph using fluent interface
func (fd *FluentDocument) AddParagraph() *ParagraphBuilder {
	return NewParagraphBuilder()
}

// AddTable adds a table using fluent interface
func (fd *FluentDocument) AddTable() *TableBuilder {
	return NewTableBuilder()
}

// Save saves the document
func (fd *FluentDocument) Save(filename string) error {
	return fd.Operations.Save(filename)
}
