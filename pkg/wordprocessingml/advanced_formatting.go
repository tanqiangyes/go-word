// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// AdvancedFormatter represents advanced formatting functionality
type AdvancedFormatter struct {
	Document *Document
}

// ComplexTable represents a complex table with advanced formatting
type ComplexTable struct {
	ID          string
	Rows        []ComplexTableRow
	Columns     []ComplexTableColumn
	Properties  TableProperties
	Borders     TableBorders
	Shading     TableShading
	Layout      TableLayout
}

// ComplexTableRow represents a row in a complex table
type ComplexTableRow struct {
	Index       int
	Cells       []ComplexTableCell
	Height      float64
	Hidden      bool
	Header      bool
	Properties  RowProperties
}

// ComplexTableColumn represents a column in a complex table
type ComplexTableColumn struct {
	Index       int
	Width       float64
	Hidden      bool
	Properties  ColumnProperties
}

// ComplexTableCell represents a cell in a complex table
type ComplexTableCell struct {
	Reference   string
	Content     CellContent
	Properties  CellProperties
	Borders     CellBorders
	Shading     CellShading
	Merged      bool
	MergeStart  string
	MergeEnd    string
}

// CellContent represents the content of a cell
type CellContent struct {
	Paragraphs  []types.Paragraph
	Tables      []ComplexTable
	Images      []Image
	Text        string
}

// TableProperties represents table properties
type TableProperties struct {
	Width           float64
	Alignment       string
	Indent          float64
	Borders         TableBorders
	Shading         TableShading
	Layout          TableLayout
	Caption         string
	Description     string
}

// TableBorders represents table borders
type TableBorders struct {
	Top     BorderSide
	Bottom  BorderSide
	Left    BorderSide
	Right   BorderSide
	InsideH BorderSide
	InsideV BorderSide
}

// BorderSide represents a border side
type BorderSide struct {
	Style   string
	Size    int
	Color   string
	Space   int
	Shadow  bool
}

// TableShading represents table shading
type TableShading struct {
	Fill      string
	Color     string
	ThemeFill string
	ThemeColor string
	Val       string
}

// TableLayout represents table layout
type TableLayout struct {
	Type        string
	Width       float64
	FixedLayout bool
}

// RowProperties represents row properties
type RowProperties struct {
	Height      float64
	Hidden      bool
	Header      bool
	CanSplit    bool
	TrHeight    float64
	TrHeightRule string
}

// ColumnProperties represents column properties
type ColumnProperties struct {
	Width       float64
	Hidden      bool
	BestFit     bool
	AutoFit     bool
}

// CellProperties represents cell properties
type CellProperties struct {
	Width       float64
	Height      float64
	VerticalAlignment string
	Shading     CellShading
	Borders     CellBorders
	Margins     CellMargins
	TextDirection string
	FitText     bool
	NoWrap      bool
}

// CellBorders represents cell borders
type CellBorders struct {
	Top     BorderSide
	Bottom  BorderSide
	Left    BorderSide
	Right   BorderSide
	InsideH BorderSide
	InsideV BorderSide
	TL2BR   BorderSide
	TR2BL   BorderSide
}

// CellShading represents cell shading
type CellShading struct {
	Fill      string
	Color     string
	ThemeFill string
	ThemeColor string
	Val       string
}

// CellMargins represents cell margins
type CellMargins struct {
	Top    float64
	Bottom float64
	Left   float64
	Right  float64
}

// Image represents an image in the document
type Image struct {
	ID          string
	Path        string
	Width       float64
	Height      float64
	AltText     string
	Title       string
	Description string
}

// HeaderFooter represents header or footer
type HeaderFooter struct {
	Type        HeaderFooterType
	Content     []types.Paragraph
	Properties  HeaderFooterProperties
}

// HeaderFooterType defines the type of header/footer
type HeaderFooterType int

const (
	// HeaderType for headers
	HeaderType HeaderFooterType = iota
	// FooterType for footers
	FooterType
	// FirstHeaderType for first page headers
	FirstHeaderType
	// FirstFooterType for first page footers
	FirstFooterType
	// EvenHeaderType for even page headers
	EvenHeaderType
	// EvenFooterType for even page footers
	EvenFooterType
)

// HeaderFooterProperties represents header/footer properties
type HeaderFooterProperties struct {
	DifferentFirst bool
	DifferentOddEven bool
	AlignWithMargins bool
	ScaleWithDoc     bool
}

// Section represents a document section
type Section struct {
	ID          string
	Properties  SectionProperties
	Headers     []HeaderFooter
	Footers     []HeaderFooter
	Content     []types.Paragraph
}

// SectionProperties represents section properties
type SectionProperties struct {
	PageSize       PageSize
	PageMargins    PageMargins
	Columns        Columns
	HeaderReference HeaderFooterReference
	FooterReference HeaderFooterReference
	PageNumbering  PageNumbering
	LineNumbering  LineNumbering
}

// PageSize represents page size
type PageSize struct {
	Width       float64
	Height      float64
	Orientation string
}

// PageMargins represents page margins
type PageMargins struct {
	Top         float64
	Bottom      float64
	Left        float64
	Right       float64
	Header      float64
	Footer      float64
	Gutter      float64
}

// Columns represents columns
type Columns struct {
	EqualWidth  bool
	Space       float64
	Separator   bool
	Column      []Column
}

// Column represents a column
type Column struct {
	Width       float64
	Space       float64
}

// HeaderFooterReference represents header/footer reference
type HeaderFooterReference struct {
	Type        HeaderFooterType
	ID          string
}

// PageNumbering represents page numbering
type PageNumbering struct {
	Start       int
	Format      string
	Restart     string
}

// LineNumbering represents line numbering
type LineNumbering struct {
	CountBy     int
	Start       int
	Distance    float64
	Restart     string
}

// NewAdvancedFormatter creates a new advanced formatter
func NewAdvancedFormatter(doc *Document) *AdvancedFormatter {
	return &AdvancedFormatter{
		Document: doc,
	}
}

// CreateComplexTable creates a complex table
func (af *AdvancedFormatter) CreateComplexTable(rows, cols int) *ComplexTable {
	table := &ComplexTable{
		ID:      fmt.Sprintf("table_%d", len(af.Document.mainPart.Content.Tables)+1),
		Rows:    make([]ComplexTableRow, rows),
		Columns: make([]ComplexTableColumn, cols),
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
		table.Rows[i] = ComplexTableRow{
			Index: i + 1,
			Cells: make([]ComplexTableCell, cols),
			Properties: RowProperties{
				Height:      20,
				CanSplit:    true,
				TrHeightRule: "auto",
			},
		}

		// 创建单元格
		for j := 0; j < cols; j++ {
			cellRef := fmt.Sprintf("%c%d", 'A'+j, i+1)
			table.Rows[i].Cells[j] = ComplexTableCell{
				Reference: cellRef,
				Content: CellContent{
					Paragraphs: []types.Paragraph{
						{
							Text:  fmt.Sprintf("单元格 %s", cellRef),
							Style: "Normal",
							Runs: []types.Run{
								{
									Text:     fmt.Sprintf("单元格 %s", cellRef),
									FontSize: 11,
									FontName: "Arial",
								},
							},
						},
					},
					Text: fmt.Sprintf("单元格 %s", cellRef),
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
		table.Columns[i] = ComplexTableColumn{
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

// AddComplexTable adds a complex table to the document
func (af *AdvancedFormatter) AddComplexTable(table *ComplexTable) error {
	if af.Document.mainPart == nil || af.Document.mainPart.Content == nil {
		return fmt.Errorf("document content is nil")
	}

	// 转换为简单表格格式
	simpleTable := types.Table{
		Rows:    make([]types.TableRow, len(table.Rows)),
		Columns: len(table.Columns),
	}

	for i, row := range table.Rows {
		simpleTable.Rows[i] = types.TableRow{
			Cells: make([]types.TableCell, len(row.Cells)),
		}

		for j, cell := range row.Cells {
			simpleTable.Rows[i].Cells[j] = types.TableCell{
				Text: cell.Content.Text,
			}
		}
	}

	af.Document.mainPart.Content.Tables = append(af.Document.mainPart.Content.Tables, simpleTable)

	return nil
}

// CreateHeader creates a header
func (af *AdvancedFormatter) CreateHeader(headerType HeaderFooterType) *HeaderFooter {
	header := &HeaderFooter{
		Type: headerType,
		Content: []types.Paragraph{
			{
				Text:  "页眉内容",
				Style: "Header",
				Runs: []types.Run{
					{
						Text:     "页眉内容",
						FontSize: 10,
						FontName: "Arial",
					},
				},
			},
		},
		Properties: HeaderFooterProperties{
			DifferentFirst:   false,
			DifferentOddEven: false,
			AlignWithMargins: true,
			ScaleWithDoc:     true,
		},
	}

	return header
}

// CreateFooter creates a footer
func (af *AdvancedFormatter) CreateFooter(footerType HeaderFooterType) *HeaderFooter {
	footer := &HeaderFooter{
		Type: footerType,
		Content: []types.Paragraph{
			{
				Text:  "页脚内容",
				Style: "Footer",
				Runs: []types.Run{
					{
						Text:     "页脚内容",
						FontSize: 10,
						FontName: "Arial",
					},
				},
			},
		},
		Properties: HeaderFooterProperties{
			DifferentFirst:   false,
			DifferentOddEven: false,
			AlignWithMargins: true,
			ScaleWithDoc:     true,
		},
	}

	return footer
}

// AddHeader adds a header to the document
func (af *AdvancedFormatter) AddHeader(header *HeaderFooter) error {
	if af.Document.mainPart == nil || af.Document.mainPart.Content == nil {
		return fmt.Errorf("document content is nil")
	}

	// 将页眉内容添加到文档
	for _, paragraph := range header.Content {
		af.Document.mainPart.Content.Paragraphs = append(af.Document.mainPart.Content.Paragraphs, paragraph)
	}

	return nil
}

// AddFooter adds a footer to the document
func (af *AdvancedFormatter) AddFooter(footer *HeaderFooter) error {
	if af.Document.mainPart == nil || af.Document.mainPart.Content == nil {
		return fmt.Errorf("document content is nil")
	}

	// 将页脚内容添加到文档
	for _, paragraph := range footer.Content {
		af.Document.mainPart.Content.Paragraphs = append(af.Document.mainPart.Content.Paragraphs, paragraph)
	}

	return nil
}

// CreateSection creates a new section
func (af *AdvancedFormatter) CreateSection() *Section {
	section := &Section{
		ID: fmt.Sprintf("section_%d", len(af.Document.mainPart.Content.Paragraphs)+1),
		Properties: SectionProperties{
			PageSize: PageSize{
				Width:       612, // 8.5 inches
				Height:      792, // 11 inches
				Orientation: "portrait",
			},
			PageMargins: PageMargins{
				Top:    72, Bottom: 72,
				Left:   72, Right:  72,
				Header: 36, Footer: 36,
			},
			Columns: Columns{
				EqualWidth: true,
				Space:      720,
			},
			PageNumbering: PageNumbering{
				Start:  1,
				Format: "decimal",
			},
		},
		Headers: make([]HeaderFooter, 0),
		Footers: make([]HeaderFooter, 0),
		Content: make([]types.Paragraph, 0),
	}

	return section
}

// AddSection adds a section to the document
func (af *AdvancedFormatter) AddSection(section *Section) error {
	if af.Document.mainPart == nil || af.Document.mainPart.Content == nil {
		return fmt.Errorf("document content is nil")
	}

	// 添加分页符
	pageBreakParagraph := types.Paragraph{
		Text:  "",
		Style: "PageBreak",
		Runs:  []types.Run{},
	}

	af.Document.mainPart.Content.Paragraphs = append(af.Document.mainPart.Content.Paragraphs, pageBreakParagraph)

	// 添加节内容
	for _, paragraph := range section.Content {
		af.Document.mainPart.Content.Paragraphs = append(af.Document.mainPart.Content.Paragraphs, paragraph)
	}

	return nil
}

// AddPageBreak adds a page break
func (af *AdvancedFormatter) AddPageBreak() error {
	if af.Document.mainPart == nil || af.Document.mainPart.Content == nil {
		return fmt.Errorf("document content is nil")
	}

	pageBreakParagraph := types.Paragraph{
		Text:  "",
		Style: "PageBreak",
		Runs:  []types.Run{},
	}

	af.Document.mainPart.Content.Paragraphs = append(af.Document.mainPart.Content.Paragraphs, pageBreakParagraph)

	return nil
}

// MergeCells merges cells in a table
func (af *AdvancedFormatter) MergeCells(table *ComplexTable, startRef, endRef string) error {
	startCol, startRow, err := parseCellReference(startRef)
	if err != nil {
		return fmt.Errorf("invalid start reference: %w", err)
	}

	endCol, endRow, err := parseCellReference(endRef)
	if err != nil {
		return fmt.Errorf("invalid end reference: %w", err)
	}

	// 查找并合并单元格
	for i := startRow - 1; i < endRow; i++ {
		if i >= len(table.Rows) {
			break
		}

		for j := startCol - 1; j < endCol; j++ {
			if j >= len(table.Rows[i].Cells) {
				break
			}

			// 设置合并属性
			table.Rows[i].Cells[j].Merged = true
			table.Rows[i].Cells[j].MergeStart = startRef
			table.Rows[i].Cells[j].MergeEnd = endRef

			// 只保留第一个单元格的内容
			if i == startRow-1 && j == startCol-1 {
				table.Rows[i].Cells[j].Merged = false
			} else {
				table.Rows[i].Cells[j].Content.Text = ""
				table.Rows[i].Cells[j].Content.Paragraphs = []types.Paragraph{}
			}
		}
	}

	return nil
}

// SetCellBorders sets cell borders
func (af *AdvancedFormatter) SetCellBorders(table *ComplexTable, cellRef string, borders CellBorders) error {
	col, row, err := parseCellReference(cellRef)
	if err != nil {
		return fmt.Errorf("invalid cell reference: %w", err)
	}

	if row-1 >= len(table.Rows) || col-1 >= len(table.Rows[row-1].Cells) {
		return fmt.Errorf("cell reference out of range")
	}

	table.Rows[row-1].Cells[col-1].Borders = borders

	return nil
}

// SetCellShading sets cell shading
func (af *AdvancedFormatter) SetCellShading(table *ComplexTable, cellRef string, shading CellShading) error {
	col, row, err := parseCellReference(cellRef)
	if err != nil {
		return fmt.Errorf("invalid cell reference: %w", err)
	}

	if row-1 >= len(table.Rows) || col-1 >= len(table.Rows[row-1].Cells) {
		return fmt.Errorf("cell reference out of range")
	}

	table.Rows[row-1].Cells[col-1].Shading = shading

	return nil
}

// parseCellReference parses a cell reference (e.g., "A1", "B2")
func parseCellReference(ref string) (col int, row int, err error) {
	// 分离列和行
	colStr := ""
	rowStr := ""
	
	for _, char := range ref {
		if char >= 'A' && char <= 'Z' {
			colStr += string(char)
		} else if char >= '0' && char <= '9' {
			rowStr += string(char)
		}
	}

	if colStr == "" || rowStr == "" {
		return 0, 0, fmt.Errorf("invalid cell reference: %s", ref)
	}

	// 转换列为数字
	col = 0
	for _, char := range colStr {
		col = col*26 + int(char-'A'+1)
	}

	// 转换行为数字
	fmt.Sscanf(rowStr, "%d", &row)

	return col, row, nil
} 