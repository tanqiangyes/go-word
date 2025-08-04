// Package parser provides specialized parsing for WordprocessingML documents
package parser

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	
	"github.com/tanqiangyes/go-word/pkg/types"
)

// WordMLParser provides WordprocessingML specific parsing
type WordMLParser struct{}

// WordDocument represents the complete Word document structure
type WordDocument struct {
	XMLName xml.Name `xml:"w:document"`
	Body    WordBody `xml:"w:body"`
}

// WordBody represents the document body
type WordBody struct {
	XMLName    xml.Name        `xml:"w:body"`
	Paragraphs []WordParagraph `xml:"w:p"`
	Tables     []WordTable     `xml:"w:tbl"`
}

// WordParagraph represents a paragraph in Word
type WordParagraph struct {
	XMLName    xml.Name        `xml:"w:p"`
	Properties *ParagraphProps `xml:"w:pPr,omitempty"`
	Runs       []WordRun      `xml:"w:r"`
	Text       string         `xml:",chardata"`
}

// ParagraphProps represents paragraph properties
type ParagraphProps struct {
	XMLName xml.Name `xml:"w:pPr"`
	Style   *Style   `xml:"w:pStyle,omitempty"`
}

// Style represents a style reference
type Style struct {
	XMLName xml.Name `xml:"w:pStyle"`
	Val     string   `xml:"w:val,attr"`
}

// WordRun represents a text run
type WordRun struct {
	XMLName    xml.Name     `xml:"w:r"`
	Properties *RunProps    `xml:"w:rPr,omitempty"`
	Text       *WordText    `xml:"w:t,omitempty"`
	Tab        *Tab         `xml:"w:tab,omitempty"`
	Break      *Break       `xml:"w:br,omitempty"`
}

// RunProps represents run properties
type RunProps struct {
	XMLName    xml.Name `xml:"w:rPr"`
	Bold       *types.Bold    `xml:"w:b,omitempty"`
	Italic     *types.Italic  `xml:"w:i,omitempty"`
	Underline  *types.Underline `xml:"w:u,omitempty"`
	Size       *types.Size    `xml:"w:sz,omitempty"`
	Font       *types.Font    `xml:"w:rFonts,omitempty"`
	Color      *types.Color   `xml:"w:color,omitempty"`
}

// WordText represents text content
type WordText struct {
	XMLName xml.Name `xml:"w:t"`
	Content string   `xml:",chardata"`
	Space   string   `xml:"xml:space,attr,omitempty"`
}

// Tab represents a tab character
type Tab struct {
	XMLName xml.Name `xml:"w:tab"`
}

// Break represents a line break
type Break struct {
	XMLName xml.Name `xml:"w:br"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

// Color represents text color
type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// WordTable represents a table
type WordTable struct {
	XMLName    xml.Name        `xml:"w:tbl"`
	Properties *TableProps     `xml:"w:tblPr,omitempty"`
	Rows       []WordTableRow  `xml:"w:tr"`
}

// TableProps represents table properties
type TableProps struct {
	XMLName xml.Name `xml:"w:tblPr"`
	Style   *Style   `xml:"w:tblStyle,omitempty"`
}

// WordTableRow represents a table row
type WordTableRow struct {
	XMLName    xml.Name         `xml:"w:tr"`
	Properties *RowProps        `xml:"w:trPr,omitempty"`
	Cells      []WordTableCell `xml:"w:tc"`
}

// RowProps represents row properties
type RowProps struct {
	XMLName xml.Name `xml:"w:trPr"`
}

// WordTableCell represents a table cell
type WordTableCell struct {
	XMLName    xml.Name        `xml:"w:tc"`
	Properties *CellProps      `xml:"w:tcPr,omitempty"`
	Content    []interface{}   `xml:",any"`
}

// CellProps represents cell properties
type CellProps struct {
	XMLName xml.Name `xml:"w:tcPr"`
}

// ParseWordDocument parses a Word document XML
func (p *WordMLParser) ParseWordDocument(content []byte) (*WordDocument, error) {
	var doc WordDocument
	if err := xml.Unmarshal(content, &doc); err != nil {
		return nil, fmt.Errorf("failed to parse Word document: %w", err)
	}
	
	return &doc, nil
}

// ExtractText extracts plain text from the document
func (p *WordMLParser) ExtractText(doc *WordDocument) string {
	var text strings.Builder
	
	for _, paragraph := range doc.Body.Paragraphs {
		paragraphText := p.extractParagraphText(paragraph)
		if paragraphText != "" {
			text.WriteString(paragraphText)
			text.WriteString("\n")
		}
	}
	
	return text.String()
}

// ExtractParagraphs extracts paragraphs with formatting
func (p *WordMLParser) ExtractParagraphs(doc *WordDocument) []types.Paragraph {
	var paragraphs []types.Paragraph

	for _, wp := range doc.Body.Paragraphs {
		paragraph := types.Paragraph{
			Runs: make([]types.Run, 0, len(wp.Runs)),
		}

		// Extract style information
		if wp.Properties != nil && wp.Properties.Style != nil {
			paragraph.Style = wp.Properties.Style.Val
		}

		// Extract runs
		for _, run := range wp.Runs {
			wordRun := p.convertRun(run)
			paragraph.Runs = append(paragraph.Runs, wordRun)
			paragraph.Text += wordRun.Text
		}

		paragraphs = append(paragraphs, paragraph)
	}

	return paragraphs
}

// ExtractTables extracts tables from the document
func (p *WordMLParser) ExtractTables(doc *WordDocument) []types.Table {
	var tables []types.Table

	for _, wt := range doc.Body.Tables {
		table := types.Table{
			Rows: make([]types.TableRow, 0, len(wt.Rows)),
		}

		for _, row := range wt.Rows {
			tableRow := types.TableRow{
				Cells: make([]types.TableCell, 0, len(row.Cells)),
			}

			for _, cell := range row.Cells {
				cellText := p.extractCellText(cell)
				tableCell := types.TableCell{
					Text: cellText,
				}
				tableRow.Cells = append(tableRow.Cells, tableCell)
			}

			table.Rows = append(table.Rows, tableRow)
		}

		if len(table.Rows) > 0 {
			table.Columns = len(table.Rows[0].Cells)
		}

		tables = append(tables, table)
	}

	return tables
}

// extractParagraphText extracts text from a paragraph
func (p *WordMLParser) extractParagraphText(paragraph WordParagraph) string {
	var text strings.Builder
	
	for _, run := range paragraph.Runs {
		if run.Text != nil {
			text.WriteString(run.Text.Content)
		}
		if run.Tab != nil {
			text.WriteString("\t")
		}
		if run.Break != nil {
			text.WriteString("\n")
		}
	}
	
	return text.String()
}

// convertRun converts a WordRun to types.Run
func (p *WordMLParser) convertRun(run WordRun) types.Run {
	wordRun := types.Run{}

	if run.Text != nil {
		wordRun.Text = run.Text.Content
	}

	if run.Properties != nil {
		if run.Properties.Bold != nil {
			wordRun.Bold = run.Properties.Bold.Val != "false"
		}
		if run.Properties.Italic != nil {
			wordRun.Italic = run.Properties.Italic.Val != "false"
		}
		if run.Properties.Underline != nil {
			wordRun.Underline = run.Properties.Underline.Val != "none"
		}
		if run.Properties.Size != nil {
			if size, err := strconv.Atoi(run.Properties.Size.Val); err == nil {
				wordRun.FontSize = size
			}
		}
		if run.Properties.Font != nil {
			wordRun.FontName = run.Properties.Font.Ascii
			if wordRun.FontName == "" {
				wordRun.FontName = run.Properties.Font.HAnsi
			}
		}
		if run.Properties.Color != nil {
			wordRun.Color = run.Properties.Color.Val
		}
	}

	return wordRun
}

// extractCellText extracts text from a table cell
func (p *WordMLParser) extractCellText(cell WordTableCell) string {
	var text strings.Builder
	
	// This is a simplified implementation
	// In a full implementation, we would need to handle various XML elements
	// that can appear in table cells (paragraphs, runs, etc.)
	
	return text.String()
} 