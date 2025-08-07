// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"time"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// DocumentBuilder provides a fluent interface for building documents
type DocumentBuilder struct {
	document *Document
	config   *DocumentConfig
}

// DocumentConfig holds configuration options for document creation
type DocumentConfig struct {
	Title       string
	Author      string
	Subject     string
	Keywords    []string
	Category    string
	Comments    string
	Language    string
	Template    string
	Protection  *ProtectionConfig
	Formatting  *FormattingConfig
	Validation  *ValidationConfig
}

// ProtectionConfig holds document protection settings
type ProtectionConfig struct {
	Enabled       bool
	Password      string
	ProtectionType ProtectionType
	Permissions   map[string]bool
	Watermark     *WatermarkConfig
}

// WatermarkConfig holds watermark settings
type WatermarkConfig struct {
	Text        string
	Font        string
	Size        int
	Color       string
	Transparency float64
	Rotation    float64
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
	Enabled     bool
	Rules       []ValidationRule
	AutoFix     bool
	StrictMode  bool
}

// NewDocumentBuilder creates a new document builder with fluent interface
func NewDocumentBuilder() *DocumentBuilder {
	return &DocumentBuilder{
		config: &DocumentConfig{
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
			Protection: &ProtectionConfig{
				Enabled: false,
			},
			Validation: &ValidationConfig{
				Enabled: true,
				AutoFix: false,
			},
		},
	}
}

// WithTitle sets the document title
func (b *DocumentBuilder) WithTitle(title string) *DocumentBuilder {
	b.config.Title = title
	return b
}

// WithAuthor sets the document author
func (b *DocumentBuilder) WithAuthor(author string) *DocumentBuilder {
	b.config.Author = author
	return b
}

// WithSubject sets the document subject
func (b *DocumentBuilder) WithSubject(subject string) *DocumentBuilder {
	b.config.Subject = subject
	return b
}

// WithKeywords sets the document keywords
func (b *DocumentBuilder) WithKeywords(keywords ...string) *DocumentBuilder {
	b.config.Keywords = keywords
	return b
}

// WithLanguage sets the document language
func (b *DocumentBuilder) WithLanguage(language string) *DocumentBuilder {
	b.config.Language = language
	return b
}

// WithTemplate sets the document template
func (b *DocumentBuilder) WithTemplate(template string) *DocumentBuilder {
	b.config.Template = template
	return b
}

// WithProtection enables document protection
func (b *DocumentBuilder) WithProtection(protectionType ProtectionType, password string) *DocumentBuilder {
	b.config.Protection.Enabled = true
	b.config.Protection.ProtectionType = protectionType
	b.config.Protection.Password = password
	return b
}

// WithPermissions sets document permissions
func (b *DocumentBuilder) WithPermissions(permissions map[string]bool) *DocumentBuilder {
	b.config.Protection.Permissions = permissions
	return b
}

// WithWatermark adds a watermark to the document
func (b *DocumentBuilder) WithWatermark(text, font string, size int, color string) *DocumentBuilder {
	b.config.Protection.Watermark = &WatermarkConfig{
		Text:        text,
		Font:        font,
		Size:        size,
		Color:       color,
		Transparency: 0.5,
		Rotation:    45,
	}
	return b
}

// WithDefaultFont sets the default font
func (b *DocumentBuilder) WithDefaultFont(font string, size int) *DocumentBuilder {
	b.config.Formatting.DefaultFont = font
	b.config.Formatting.DefaultFontSize = size
	return b
}

// WithMargins sets the page margins
func (b *DocumentBuilder) WithMargins(top, bottom, left, right float64) *DocumentBuilder {
	b.config.Formatting.Margins.Top = top
	b.config.Formatting.Margins.Bottom = bottom
	b.config.Formatting.Margins.Left = left
	b.config.Formatting.Margins.Right = right
	return b
}

// WithPageSize sets the page size
func (b *DocumentBuilder) WithPageSize(width, height float64, orientation string) *DocumentBuilder {
	b.config.Formatting.PageSize.Width = width
	b.config.Formatting.PageSize.Height = height
	b.config.Formatting.PageSize.Orientation = orientation
	return b
}

// WithTheme sets the document theme
func (b *DocumentBuilder) WithTheme(theme string) *DocumentBuilder {
	b.config.Formatting.Theme = theme
	return b
}

// WithValidation enables document validation
func (b *DocumentBuilder) WithValidation(enabled, autoFix, strictMode bool) *DocumentBuilder {
	b.config.Validation.Enabled = enabled
	b.config.Validation.AutoFix = autoFix
	b.config.Validation.StrictMode = strictMode
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
	
	b.document = doc
	return doc, nil
}

// applyConfiguration applies the builder configuration to the document
func (b *DocumentBuilder) applyConfiguration(doc *Document) error {
	// Set document properties
	if b.config.Title != "" {
		// TODO: Set document title property
	}
	
	if b.config.Author != "" {
		// TODO: Set document author property
	}
	
	// Apply protection if enabled
	if b.config.Protection.Enabled {
		// TODO: Implement protection
	}
	
	// Apply validation if enabled
	if b.config.Validation.Enabled {
		// TODO: Implement validation
	}
	
	return nil
}

// ParagraphBuilder provides a fluent interface for building paragraphs
type ParagraphBuilder struct {
	paragraph types.Paragraph
	content   []types.Run
}

// NewParagraphBuilder creates a new paragraph builder
func NewParagraphBuilder() *ParagraphBuilder {
	return &ParagraphBuilder{
		paragraph: types.Paragraph{
			Runs: make([]types.Run, 0),
		},
	}
}

// WithText adds plain text to the paragraph
func (b *ParagraphBuilder) WithText(text string) *ParagraphBuilder {
	b.paragraph.Text = text
	b.paragraph.Runs = append(b.paragraph.Runs, types.Run{
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
	
	b.paragraph.Runs = append(b.paragraph.Runs, run)
	return b
}

// WithStyle sets the paragraph style
func (b *ParagraphBuilder) WithStyle(style string) *ParagraphBuilder {
	b.paragraph.Style = style
	return b
}

// WithComment adds a comment to the paragraph
func (b *ParagraphBuilder) WithComment(author, text string) *ParagraphBuilder {
	b.paragraph.HasComment = true
	b.paragraph.CommentID = fmt.Sprintf("comment_%d", time.Now().Unix())
	// TODO: Add comment to comment manager
	return b
}

// Build creates the paragraph
func (b *ParagraphBuilder) Build() types.Paragraph {
	return b.paragraph
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
	table   types.Table
	rows    [][]string
	headers []string
	style   string
}

// NewTableBuilder creates a new table builder
func NewTableBuilder() *TableBuilder {
	return &TableBuilder{
		table: types.Table{
			Rows: make([]types.TableRow, 0),
		},
	}
}

// WithHeaders sets the table headers
func (b *TableBuilder) WithHeaders(headers ...string) *TableBuilder {
	b.headers = headers
	b.table.Columns = len(headers)
	return b
}

// WithRows adds rows to the table
func (b *TableBuilder) WithRows(rows ...[]string) *TableBuilder {
	b.rows = append(b.rows, rows...)
	return b
}

// WithStyle sets the table style
func (b *TableBuilder) WithStyle(style string) *TableBuilder {
	b.style = style
	return b
}

// Build creates the table
func (b *TableBuilder) Build() types.Table {
	// Add header row if headers are provided
	if len(b.headers) > 0 {
		headerRow := types.TableRow{
			Cells: make([]types.TableCell, len(b.headers)),
		}
		for i, header := range b.headers {
			headerRow.Cells[i] = types.TableCell{
				Text: header,
			}
		}
		b.table.Rows = append(b.table.Rows, headerRow)
	}
	
	// Add data rows
	for _, rowData := range b.rows {
		row := types.TableRow{
			Cells: make([]types.TableCell, len(rowData)),
		}
		for i, cellData := range rowData {
			row.Cells[i] = types.TableCell{
				Text: cellData,
			}
		}
		b.table.Rows = append(b.table.Rows, row)
	}
	
	return b.table
}

// DocumentOperations provides a fluent interface for document operations
type DocumentOperations struct {
	document *Document
}

// NewDocumentOperations creates document operations for a document
func NewDocumentOperations(doc *Document) *DocumentOperations {
	return &DocumentOperations{
		document: doc,
	}
}

// AddParagraph adds a paragraph to the document
func (ops *DocumentOperations) AddParagraph(builder *ParagraphBuilder) *DocumentOperations {
	_ = builder.Build()
	// TODO: Add paragraph to document
	return ops
}

// AddTable adds a table to the document
func (ops *DocumentOperations) AddTable(builder *TableBuilder) *DocumentOperations {
	_ = builder.Build()
	// TODO: Add table to document
	return ops
}

// Save saves the document to a file
func (ops *DocumentOperations) Save(filename string) error {
	// TODO: Implement save functionality
	return nil
}

// FluentDocument provides a fluent interface for document operations
type FluentDocument struct {
	document *Document
	operations *DocumentOperations
}

// NewFluentDocument creates a new fluent document interface
func NewFluentDocument(doc *Document) *FluentDocument {
	return &FluentDocument{
		document: doc,
		operations: NewDocumentOperations(doc),
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
	return fd.operations.Save(filename)
} 