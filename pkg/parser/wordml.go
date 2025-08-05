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
	XMLName xml.Name `xml:"document"`
	Body    WordBody `xml:"body"`
}

// WordBody represents the document body
type WordBody struct {
	XMLName    xml.Name        `xml:"body"`
	Paragraphs []WordParagraph `xml:"p"`
	Tables     []WordTable     `xml:"tbl"`
}

// WordParagraph represents a paragraph in Word
type WordParagraph struct {
	XMLName    xml.Name        `xml:"p"`
	Properties *ParagraphProps `xml:"pPr,omitempty"`
	Runs       []WordRun      `xml:"r"`
	Text       string         `xml:",chardata"`
}

// ParagraphProps represents paragraph properties
type ParagraphProps struct {
	XMLName xml.Name `xml:"pPr"`
	Style   *Style   `xml:"pStyle,omitempty"`
}

// Style represents a style reference
type Style struct {
	XMLName xml.Name `xml:"pStyle"`
	Val     string   `xml:"val,attr"`
}

// WordRun represents a text run
type WordRun struct {
	XMLName    xml.Name     `xml:"r"`
	Properties *RunProps    `xml:"rPr,omitempty"`
	Text       *WordText    `xml:"t,omitempty"`
	Tab        *Tab         `xml:"tab,omitempty"`
	Break      *Break       `xml:"br,omitempty"`
}

// RunProps represents run properties
type RunProps struct {
	XMLName    xml.Name `xml:"rPr"`
	Bold       *types.Bold    `xml:"b,omitempty"`
	Italic     *types.Italic  `xml:"i,omitempty"`
	Underline  *types.Underline `xml:"u,omitempty"`
	Size       *types.Size    `xml:"sz,omitempty"`
	Font       *types.Font    `xml:"rFonts,omitempty"`
	Color      *types.Color   `xml:"color,omitempty"`
}

// WordText represents text content
type WordText struct {
	XMLName xml.Name `xml:"t"`
	Content string   `xml:",chardata"`
	Space   string   `xml:"space,attr,omitempty"`
}

// Tab represents a tab character
type Tab struct {
	XMLName xml.Name `xml:"tab"`
}

// Break represents a line break
type Break struct {
	XMLName xml.Name `xml:"br"`
	Type    string   `xml:"type,attr,omitempty"`
}

// Color represents text color
type Color struct {
	XMLName xml.Name `xml:"color"`
	Val     string   `xml:"val,attr,omitempty"`
}

// WordTable represents a table
type WordTable struct {
	XMLName    xml.Name        `xml:"tbl"`
	Properties *TableProps     `xml:"tblPr,omitempty"`
	Rows       []WordTableRow  `xml:"tr"`
}

// TableProps represents table properties
type TableProps struct {
	XMLName xml.Name `xml:"tblPr"`
	Style   *Style   `xml:"tblStyle,omitempty"`
}

// WordTableRow represents a table row
type WordTableRow struct {
	XMLName    xml.Name         `xml:"tr"`
	Properties *RowProps        `xml:"trPr,omitempty"`
	Cells      []WordTableCell `xml:"tc"`
}

// RowProps represents row properties
type RowProps struct {
	XMLName xml.Name `xml:"trPr"`
}

// WordTableCell represents a table cell
type WordTableCell struct {
	XMLName    xml.Name        `xml:"tc"`
	Properties *CellProps      `xml:"tcPr,omitempty"`
	Paragraphs []WordParagraph `xml:"p"`
}

// CellProps represents cell properties
type CellProps struct {
	XMLName xml.Name `xml:"tcPr"`
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
	
	// 遍历单元格内容，查找段落和文本
	for _, content := range cell.Paragraphs {
		text.WriteString(p.extractParagraphText(content))
	}
	
	return text.String()
} 