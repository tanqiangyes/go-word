// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// FormatSupport represents format support functionality
type FormatSupport struct {
	Document *Document
}

// RichTextContent represents rich text content
type RichTextContent struct {
	Text        string
	Formatting  RichTextFormatting
	Hyperlinks  []Hyperlink
	Images      []Image
	Tables      []RichTextTable
	Lists       []RichTextList
}

// RichTextFormatting represents rich text formatting
type RichTextFormatting struct {
	Font        Font
	Paragraph   ParagraphFormat
	Borders     BorderFormat
	Shading     ShadingFormat
	Effects     TextEffects
}

// Font represents font properties
type Font struct {
	Name      string
	Size      float64
	Bold      bool
	Italic    bool
	Underline bool
	Strike    bool
	Color     string
	Highlight string
	Subscript bool
	Superscript bool
	SmallCaps bool
	AllCaps   bool
}

// ParagraphFormat represents paragraph formatting
type ParagraphFormat struct {
	Alignment      string
	Indent         IndentFormat
	Spacing        SpacingFormat
	Borders        BorderFormat
	Shading        ShadingFormat
	KeepLines      bool
	KeepNext       bool
	PageBreakBefore bool
	WidowControl   bool
}

// IndentFormat represents indentation formatting
type IndentFormat struct {
	Left    float64
	Right   float64
	First   float64
	Hanging float64
}

// SpacingFormat represents spacing formatting
type SpacingFormat struct {
	Before  float64
	After   float64
	Line    float64
	Between bool
}

// BorderFormat represents border formatting
type BorderFormat struct {
	Top     BorderSide
	Bottom  BorderSide
	Left    BorderSide
	Right   BorderSide
	Between BorderSide
	Bar     BorderSide
}

// ShadingFormat represents shading formatting
type ShadingFormat struct {
	Fill      string
	Color     string
	ThemeFill string
	ThemeColor string
	Val       string
}

// TextEffects represents text effects
type TextEffects struct {
	Shadow     bool
	Outline    bool
	Emboss     bool
	Engrave    bool
	Reflection bool
	Glow       bool
	SoftEdge   bool
}

// Hyperlink represents a hyperlink
type Hyperlink struct {
	URL         string
	Text        string
	Tooltip     string
	Target      string
	Style       string
}

// RichTextTable represents a rich text table
type RichTextTable struct {
	Rows        []RichTextRow
	Columns     []RichTextColumn
	Properties  TableProperties
	Borders     TableBorders
	Shading     TableShading
}

// RichTextRow represents a row in rich text table
type RichTextRow struct {
	Index       int
	Cells       []RichTextCell
	Height      float64
	Hidden      bool
	Header      bool
	Properties  RowProperties
}

// RichTextColumn represents a column in rich text table
type RichTextColumn struct {
	Index       int
	Width       float64
	Hidden      bool
	Properties  ColumnProperties
}

// RichTextCell represents a cell in rich text table
type RichTableCell struct {
	Reference   string
	Content     RichTextContent
	Properties  CellProperties
	Borders     CellBorders
	Shading     CellShading
	Merged      bool
	MergeStart  string
	MergeEnd    string
}

// RichTextList represents a rich text list
type RichTextList struct {
	ID          string
	Type        ListType
	Level       int
	Items       []RichTextListItem
	Properties  ListProperties
}

// RichTextListItem represents a list item
type RichTextListItem struct {
	Index       int
	Content     RichTextContent
	Level       int
	Properties  ListItemProperties
}

// ListType defines the type of list
type ListType int

const (
	// BulletList for bullet lists
	BulletList ListType = iota
	// NumberedList for numbered lists
	NumberedList
	// CustomList for custom lists
	CustomList
)

// ListProperties represents list properties
type ListProperties struct {
	Type        ListType
	Start       int
	Restart     bool
	Level       int
	Format      string
	Indent      float64
	Hanging     float64
}

// ListItemProperties represents list item properties
type ListItemProperties struct {
	Level       int
	Number      int
	Format      string
	Indent      float64
	Hanging     float64
}

// DocumentFormat represents document format
type DocumentFormat int

const (
	// DocxFormat for .docx files
	DocxFormat DocumentFormat = iota
	// DocFormat for .doc files
	DocFormat
	// DocmFormat for .docm files
	DocmFormat
	// RtfFormat for .rtf files
	RtfFormat
)

// NewFormatSupport creates a new format support
func NewFormatSupport(doc *Document) *FormatSupport {
	return &FormatSupport{
		Document: doc,
	}
}

// DetectFormat detects the format of a document
func (fs *FormatSupport) DetectFormat(filename string) (DocumentFormat, error) {
	ext := strings.ToLower(getFileExtension(filename))
	
	switch ext {
	case ".docx":
		return DocxFormat, nil
	case ".doc":
		return DocFormat, nil
	case ".docm":
		return DocmFormat, nil
	case ".rtf":
		return RtfFormat, nil
	default:
		return DocxFormat, fmt.Errorf("unsupported format: %s", ext)
	}
}

// ConvertFormat converts document to a different format
func (fs *FormatSupport) ConvertFormat(targetFormat DocumentFormat) error {
	switch targetFormat {
	case DocxFormat:
		return fs.convertToDocx()
	case DocFormat:
		return fs.convertToDoc()
	case DocmFormat:
		return fs.convertToDocm()
	case RtfFormat:
		return fs.convertToRtf()
	default:
		return fmt.Errorf("unsupported target format: %v", targetFormat)
	}
}

// convertToDocx converts to .docx format
func (fs *FormatSupport) convertToDocx() error {
	// .docx is the native format, no conversion needed
	return nil
}

// convertToDoc converts to .doc format
func (fs *FormatSupport) convertToDoc() error {
	// 实现.doc格式转换
	// 这需要处理旧的Word二进制格式
	return fmt.Errorf("conversion to .doc format not yet implemented")
}

// convertToDocm converts to .docm format
func (fs *FormatSupport) convertToDocm() error {
	// .docm格式是包含宏的.docx格式
	// 需要添加宏支持
	return fmt.Errorf("conversion to .docm format not yet implemented")
}

// convertToRtf converts to .rtf format
func (fs *FormatSupport) convertToRtf() error {
	// 实现RTF格式转换
	return fmt.Errorf("conversion to .rtf format not yet implemented")
}

// CreateRichTextContent creates rich text content
func (fs *FormatSupport) CreateRichTextContent(text string) *RichTextContent {
	return &RichTextContent{
		Text: text,
		Formatting: RichTextFormatting{
			Font: Font{
				Name: "Arial",
				Size: 11,
			},
			Paragraph: ParagraphFormat{
				Alignment: "left",
				Indent: IndentFormat{
					Left: 0,
					Right: 0,
					First: 0,
				},
				Spacing: SpacingFormat{
					Before: 0,
					After: 0,
					Line: 1.0,
				},
			},
		},
		Hyperlinks: make([]Hyperlink, 0),
		Images:     make([]Image, 0),
		Tables:     make([]RichTextTable, 0),
		Lists:      make([]RichTextList, 0),
	}
}

// AddRichTextFormatting adds rich text formatting
func (fs *FormatSupport) AddRichTextFormatting(content *RichTextContent, formatting RichTextFormatting) {
	content.Formatting = formatting
}

// AddHyperlink adds a hyperlink to rich text content
func (fs *FormatSupport) AddHyperlink(content *RichTextContent, url, text, tooltip string) {
	hyperlink := Hyperlink{
		URL:     url,
		Text:    text,
		Tooltip: tooltip,
		Target:  "_blank",
		Style:   "hyperlink",
	}
	
	content.Hyperlinks = append(content.Hyperlinks, hyperlink)
}

// AddImage adds an image to rich text content
func (fs *FormatSupport) AddImage(content *RichTextContent, path string, width, height float64) {
	image := Image{
		ID:     fmt.Sprintf("image_%d", len(content.Images)+1),
		Path:   path,
		Width:  width,
		Height: height,
		AltText: "图片",
		Title:   "图片标题",
	}
	
	content.Images = append(content.Images, image)
}

// CreateRichTextTable creates a rich text table
func (fs *FormatSupport) CreateRichTextTable(rows, cols int) *RichTextTable {
	table := &RichTextTable{
		Rows:    make([]RichTextRow, rows),
		Columns: make([]RichTextColumn, cols),
		Properties: TableProperties{
			Width:     100,
			Alignment: "left",
			Layout: TableLayout{
				Type:        "fixed",
				Width:       100,
				FixedLayout: true,
			},
		},
		Borders: TableBorders{
			Top:     BorderSide{Style: "single", Size: 1, Color: "000000"},
			Bottom:  BorderSide{Style: "single", Size: 1, Color: "000000"},
			Left:    BorderSide{Style: "single", Size: 1, Color: "000000"},
			Right:   BorderSide{Style: "single", Size: 1, Color: "000000"},
			InsideH: BorderSide{Style: "single", Size: 1, Color: "000000"},
			InsideV: BorderSide{Style: "single", Size: 1, Color: "000000"},
		},
	}

	// 创建行
	for i := 0; i < rows; i++ {
		table.Rows[i] = RichTextRow{
			Index: i + 1,
			Cells: make([]RichTextCell, cols),
			Properties: RowProperties{
				Height:      20,
				CanSplit:    true,
				TrHeightRule: "auto",
			},
		}

		// 创建单元格
		for j := 0; j < cols; j++ {
			cellRef := fmt.Sprintf("%c%d", 'A'+j, i+1)
			table.Rows[i].Cells[j] = RichTextCell{
				Reference: cellRef,
				Content: RichTextContent{
					Text: fmt.Sprintf("单元格 %s", cellRef),
					Formatting: RichTextFormatting{
						Font: Font{
							Name: "Arial",
							Size: 11,
						},
					},
				},
				Properties: CellProperties{
					Width:            20,
					Height:           20,
					VerticalAlignment: "top",
					Margins: CellMargins{
						Top:    2, Bottom: 2,
						Left:   2, Right:  2,
					},
				},
			}
		}
	}

	// 创建列
	for i := 0; i < cols; i++ {
		table.Columns[i] = RichTextColumn{
			Index: i + 1,
			Width: 20,
			Properties: ColumnProperties{
				Width:   20,
				BestFit: true,
			},
		}
	}

	return table
}

// CreateRichTextList creates a rich text list
func (fs *FormatSupport) CreateRichTextList(listType ListType) *RichTextList {
	list := &RichTextList{
		ID:   fmt.Sprintf("list_%d", len(fs.Document.mainPart.Content.Paragraphs)+1),
		Type: listType,
		Level: 0,
		Items: make([]RichTextListItem, 0),
		Properties: ListProperties{
			Type:   listType,
			Start:  1,
			Level:  0,
			Format: getListFormat(listType),
		},
	}

	return list
}

// AddListItem adds an item to a rich text list
func (fs *FormatSupport) AddListItem(list *RichTextList, content RichTextContent, level int) {
	item := RichTextListItem{
		Index:   len(list.Items) + 1,
		Content: content,
		Level:   level,
		Properties: ListItemProperties{
			Level:  level,
			Number: len(list.Items) + 1,
			Format: getListFormat(list.Type),
		},
	}

	list.Items = append(list.Items, item)
}

// getListFormat returns the format string for a list type
func getListFormat(listType ListType) string {
	switch listType {
	case BulletList:
		return "•"
	case NumberedList:
		return "1."
	case CustomList:
		return "-"
	default:
		return "•"
	}
}

// getFileExtension gets the file extension
func getFileExtension(filename string) string {
	lastDot := strings.LastIndex(filename, ".")
	if lastDot == -1 {
		return ""
	}
	return filename[lastDot:]
}

// ApplyRichTextFormatting applies rich text formatting to a paragraph
func (fs *FormatSupport) ApplyRichTextFormatting(paragraph *types.Paragraph, formatting RichTextFormatting) {
	// 应用字体格式
	for i := range paragraph.Runs {
		run := &paragraph.Runs[i]
		run.FontName = formatting.Font.Name
		run.FontSize = int(formatting.Font.Size)
		run.Bold = formatting.Font.Bold
		run.Italic = formatting.Font.Italic
		run.Underline = formatting.Font.Underline
	}

	// 应用段落格式
	paragraph.Style = getParagraphStyle(formatting.Paragraph)
}

// getParagraphStyle returns the paragraph style based on formatting
func getParagraphStyle(format ParagraphFormat) string {
	switch format.Alignment {
	case "center":
		return "Center"
	case "right":
		return "Right"
	case "justify":
		return "Justify"
	default:
		return "Normal"
	}
} 