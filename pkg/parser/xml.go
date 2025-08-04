// Package parser provides XML parsing functionality for Word documents
package parser

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	
	"github.com/tanqiangyes/go-word/pkg/types"
)

// XMLParser provides XML parsing capabilities
type XMLParser struct{}

// DocumentXML represents the structure of word/document.xml
type DocumentXML struct {
	XMLName xml.Name `xml:"w:document"`
	Body    Body     `xml:"w:body"`
}

// Body represents the document body
type Body struct {
	XMLName    xml.Name     `xml:"w:body"`
	Paragraphs []Paragraph  `xml:"w:p"`
	Tables     []Table      `xml:"w:tbl"`
}

// Paragraph represents a paragraph element
type Paragraph struct {
	XMLName xml.Name `xml:"w:p"`
	Runs    []Run    `xml:"w:r"`
	Text    string   `xml:",chardata"`
}

// Run represents a text run element
type Run struct {
	XMLName    xml.Name `xml:"w:r"`
	Text       Text     `xml:"w:t"`
	Properties *RunProperties `xml:"w:rPr,omitempty"`
}

// Text represents a text element
type Text struct {
	XMLName xml.Name `xml:"w:t"`
	Content string   `xml:",chardata"`
	Space   string   `xml:"xml:space,attr,omitempty"`
}

// RunProperties represents run properties
type RunProperties struct {
	XMLName xml.Name `xml:"w:rPr"`
	Bold    *types.Bold `xml:"w:b,omitempty"`
	Italic  *types.Italic `xml:"w:i,omitempty"`
	Size    *types.Size `xml:"w:sz,omitempty"`
	Font    *types.Font `xml:"w:rFonts,omitempty"`
}

// Table represents a table element
type Table struct {
	XMLName xml.Name `xml:"w:tbl"`
	Rows    []TableRow `xml:"w:tr"`
}

// TableRow represents a table row
type TableRow struct {
	XMLName xml.Name    `xml:"w:tr"`
	Cells   []TableCell `xml:"w:tc"`
}

// TableCell represents a table cell
type TableCell struct {
	XMLName xml.Name `xml:"w:tc"`
	Content []interface{} `xml:",any"`
}

// ParseDocument parses the main document XML content
func (p *XMLParser) ParseDocument(content []byte) (*DocumentXML, error) {
	var doc DocumentXML
	if err := xml.Unmarshal(content, &doc); err != nil {
		return nil, fmt.Errorf("failed to parse document XML: %w", err)
	}
	
	return &doc, nil
}

// ExtractText extracts plain text from parsed document
func (p *XMLParser) ExtractText(doc *DocumentXML) string {
	var text strings.Builder
	
	// Extract text from paragraphs
	for _, paragraph := range doc.Body.Paragraphs {
		for _, run := range paragraph.Runs {
			if run.Text.Content != "" {
				text.WriteString(run.Text.Content)
			}
		}
		text.WriteString("\n")
	}
	
	return text.String()
}

// ExtractParagraphs extracts paragraphs with formatting information
func (p *XMLParser) ExtractParagraphs(doc *DocumentXML) []types.Paragraph {
	var paragraphs []types.Paragraph
	
	for _, p := range doc.Body.Paragraphs {
		paragraph := types.Paragraph{
			Runs: make([]types.Run, 0, len(p.Runs)),
		}
		
		for _, r := range p.Runs {
			run := types.Run{
				Text: r.Text.Content,
			}
			
			// Extract formatting properties
			if r.Properties != nil {
				if r.Properties.Bold != nil {
					run.Bold = r.Properties.Bold.Val != "false"
				}
				if r.Properties.Italic != nil {
					run.Italic = r.Properties.Italic.Val != "false"
				}
				if r.Properties.Size != nil {
					// Convert size to integer if possible
					if size, err := strconv.Atoi(r.Properties.Size.Val); err == nil {
						run.FontSize = size
					}
				}
				if r.Properties.Font != nil {
					run.FontName = r.Properties.Font.Ascii
					if run.FontName == "" {
						run.FontName = r.Properties.Font.HAnsi
					}
				}
			}
			
			paragraph.Runs = append(paragraph.Runs, run)
			paragraph.Text += run.Text
		}
		
		paragraphs = append(paragraphs, paragraph)
	}
	
	return paragraphs
}

// ExtractTables extracts tables from the document
func (p *XMLParser) ExtractTables(doc *DocumentXML) []types.Table {
	var tables []types.Table
	
	for _, t := range doc.Body.Tables {
		table := types.Table{
			Rows: make([]types.TableRow, 0, len(t.Rows)),
		}
		
		for _, row := range t.Rows {
			tableRow := types.TableRow{
				Cells: make([]types.TableCell, 0, len(row.Cells)),
			}
			
			for _, cell := range row.Cells {
				// Extract text from cell content
				cellText := extractTextFromCell(cell)
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

// extractTextFromCell extracts text content from a table cell
func extractTextFromCell(cell TableCell) string {
	var text strings.Builder
	
	// This is a simplified implementation
	// In a full implementation, we would need to handle various XML elements
	// that can appear in table cells (paragraphs, runs, etc.)
	
	return text.String()
} 