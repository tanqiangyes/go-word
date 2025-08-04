// Package writer provides document writing and modification functionality
package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strings"

	"github.com/go-word/pkg/opc"
	"github.com/go-word/pkg/wordprocessingml"
	"github.com/go-word/pkg/types"
)

// DocumentWriter provides functionality to modify and create Word documents
type DocumentWriter struct {
	container *opc.Container
	document  *wordprocessingml.Document
}

// NewDocumentWriter creates a new document writer
func NewDocumentWriter() *DocumentWriter {
	return &DocumentWriter{}
}

// OpenForModification opens an existing document for modification
func (w *DocumentWriter) OpenForModification(filename string) error {
	doc, err := wordprocessingml.Open(filename)
	if err != nil {
		return fmt.Errorf("failed to open document for modification: %w", err)
	}

	w.document = doc
	w.container = doc.GetContainer()
	return nil
}

// CreateNewDocument creates a new empty Word document
func (w *DocumentWriter) CreateNewDocument() error {
	// Create a new OPC container
	w.container = &opc.Container{}

	// Create basic document structure
	w.document = &wordprocessingml.Document{
		Container: w.container,
	}

	// Initialize with empty content
	w.document.MainPart = &wordprocessingml.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}

	return nil
}

// AddParagraph adds a new paragraph to the document
func (w *DocumentWriter) AddParagraph(text string, style string) error {
	if w.document == nil || w.document.MainPart == nil {
		return fmt.Errorf("document not initialized")
	}

	paragraph := types.Paragraph{
		Text:  text,
		Style: style,
		Runs: []types.Run{
			{
				Text: text,
			},
		},
	}

	w.document.MainPart.Content.Paragraphs = append(
		w.document.MainPart.Content.Paragraphs, paragraph)

	// Update document text
	w.document.MainPart.Content.Text += text + "\n"

	return nil
}

// AddFormattedParagraph adds a paragraph with specific formatting
func (w *DocumentWriter) AddFormattedParagraph(text string, style string, runs []types.Run) error {
	if w.document == nil || w.document.MainPart == nil {
		return fmt.Errorf("document not initialized")
	}

	paragraph := types.Paragraph{
		Text:  text,
		Style: style,
		Runs:  runs,
	}

	w.document.MainPart.Content.Paragraphs = append(
		w.document.MainPart.Content.Paragraphs, paragraph)

	// Update document text
	w.document.MainPart.Content.Text += text + "\n"

	return nil
}

// AddTable adds a new table to the document
func (w *DocumentWriter) AddTable(rows [][]string) error {
	if w.document == nil || w.document.MainPart == nil {
		return fmt.Errorf("document not initialized")
	}

	table := types.Table{
		Rows: make([]types.TableRow, len(rows)),
	}

	for i, rowData := range rows {
		row := types.TableRow{
			Cells: make([]types.TableCell, len(rowData)),
		}

		for j, cellText := range rowData {
			row.Cells[j] = types.TableCell{
				Text: cellText,
			}
		}

		table.Rows[i] = row
	}

	if len(rows) > 0 {
		table.Columns = len(rows[0])
	}

	w.document.MainPart.Content.Tables = append(
		w.document.MainPart.Content.Tables, table)

	return nil
}

// ReplaceText replaces all occurrences of old text with new text
func (w *DocumentWriter) ReplaceText(oldText, newText string) error {
	if w.document == nil || w.document.MainPart == nil {
		return fmt.Errorf("document not initialized")
	}

	// Replace in document text
	w.document.MainPart.Content.Text = strings.ReplaceAll(
		w.document.MainPart.Content.Text, oldText, newText)

	// Replace in paragraphs
	for i := range w.document.MainPart.Content.Paragraphs {
		paragraph := &w.document.MainPart.Content.Paragraphs[i]
		paragraph.Text = strings.ReplaceAll(paragraph.Text, oldText, newText)

		// Replace in runs
		for j := range paragraph.Runs {
			run := &paragraph.Runs[j]
			run.Text = strings.ReplaceAll(run.Text, oldText, newText)
		}
	}

	// Replace in table cells
	for i := range w.document.MainPart.Content.Tables {
		table := &w.document.MainPart.Content.Tables[i]
		for j := range table.Rows {
			row := &table.Rows[j]
			for k := range row.Cells {
				cell := &row.Cells[k]
				cell.Text = strings.ReplaceAll(cell.Text, oldText, newText)
			}
		}
	}

	return nil
}

// SetParagraphStyle sets the style of a specific paragraph
func (w *DocumentWriter) SetParagraphStyle(index int, style string) error {
	if w.document == nil || w.document.MainPart == nil {
		return fmt.Errorf("document not initialized")
	}

	if index < 0 || index >= len(w.document.MainPart.Content.Paragraphs) {
		return fmt.Errorf("paragraph index out of range")
	}

	w.document.MainPart.Content.Paragraphs[index].Style = style
	return nil
}

// SetRunFormatting sets formatting for a specific run in a paragraph
func (w *DocumentWriter) SetRunFormatting(paragraphIndex, runIndex int, formatting types.Run) error {
	if w.document == nil || w.document.MainPart == nil {
		return fmt.Errorf("document not initialized")
	}

	if paragraphIndex < 0 || paragraphIndex >= len(w.document.MainPart.Content.Paragraphs) {
		return fmt.Errorf("paragraph index out of range")
	}

	paragraph := &w.document.MainPart.Content.Paragraphs[paragraphIndex]
	if runIndex < 0 || runIndex >= len(paragraph.Runs) {
		return fmt.Errorf("run index out of range")
	}

	paragraph.Runs[runIndex] = formatting
	return nil
}

// Save saves the document to a file
func (w *DocumentWriter) Save(filename string) error {
	if w.document == nil {
		return fmt.Errorf("document not initialized")
	}

	// Generate XML content for the main document part
	xmlContent, err := w.generateDocumentXML()
	if err != nil {
		return fmt.Errorf("failed to generate document XML: %w", err)
	}

	// Create a new container with the modified content
	container := &opc.Container{}
	
	// Add the main document part
	mainPart := &opc.Part{
		Name:        "word/document.xml",
		Content:     xmlContent,
		ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
	}

	// TODO: Add other required parts (styles, relationships, etc.)
	// For now, we'll create a minimal document structure

	// Save the container to file
	return container.SaveToFile(filename)
}

// generateDocumentXML generates the XML content for the main document part
func (w *DocumentWriter) generateDocumentXML() ([]byte, error) {
	if w.document == nil || w.document.MainPart == nil {
		return nil, fmt.Errorf("document not initialized")
	}

	// Create the XML structure
	doc := &DocumentXML{
		XMLName: xml.Name{Local: "w:document"},
		XMLNS:   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Body: DocumentBody{
			XMLName: xml.Name{Local: "w:body"},
		},
	}

	// Add paragraphs
	for _, paragraph := range w.document.MainPart.Content.Paragraphs {
		xmlParagraph := ParagraphXML{
			XMLName: xml.Name{Local: "w:p"},
		}

		// Add paragraph properties if style is set
		if paragraph.Style != "" {
			xmlParagraph.Properties = &ParagraphPropertiesXML{
				XMLName: xml.Name{Local: "w:pPr"},
				Style: &StyleXML{
					XMLName: xml.Name{Local: "w:pStyle"},
					Val:     paragraph.Style,
				},
			}
		}

		// Add runs
		for _, run := range paragraph.Runs {
			xmlRun := RunXML{
				XMLName: xml.Name{Local: "w:r"},
			}

			// Add run properties if formatting is set
			if run.Bold || run.Italic || run.Underline || run.FontSize > 0 || run.FontName != "" {
				xmlRun.Properties = &RunPropertiesXML{
					XMLName: xml.Name{Local: "w:rPr"},
				}

				if run.Bold {
					xmlRun.Properties.Bold = &BoldXML{
						XMLName: xml.Name{Local: "w:b"},
						Val:     "true",
					}
				}

				if run.Italic {
					xmlRun.Properties.Italic = &ItalicXML{
						XMLName: xml.Name{Local: "w:i"},
						Val:     "true",
					}
				}

				if run.Underline {
					xmlRun.Properties.Underline = &UnderlineXML{
						XMLName: xml.Name{Local: "w:u"},
						Val:     "single",
					}
				}

				if run.FontSize > 0 {
					xmlRun.Properties.Size = &SizeXML{
						XMLName: xml.Name{Local: "w:sz"},
						Val:     fmt.Sprintf("%d", run.FontSize),
					}
				}

				if run.FontName != "" {
					xmlRun.Properties.Font = &FontXML{
						XMLName: xml.Name{Local: "w:rFonts"},
						Ascii:   run.FontName,
						HAnsi:   run.FontName,
					}
				}
			}

			// Add text
			xmlRun.Text = &TextXML{
				XMLName: xml.Name{Local: "w:t"},
				Content: run.Text,
			}

			xmlParagraph.Runs = append(xmlParagraph.Runs, xmlRun)
		}

		doc.Body.Paragraphs = append(doc.Body.Paragraphs, xmlParagraph)
	}

	// Add tables
	for _, table := range w.document.MainPart.Content.Tables {
		xmlTable := TableXML{
			XMLName: xml.Name{Local: "w:tbl"},
		}

		for _, row := range table.Rows {
			xmlRow := TableRowXML{
				XMLName: xml.Name{Local: "w:tr"},
			}

			for _, cell := range row.Cells {
				xmlCell := TableCellXML{
					XMLName: xml.Name{Local: "w:tc"},
					Content: []interface{}{
						ParagraphXML{
							XMLName: xml.Name{Local: "w:p"},
							Runs: []RunXML{
								{
									XMLName: xml.Name{Local: "w:r"},
									Text: &TextXML{
										XMLName: xml.Name{Local: "w:t"},
										Content: cell.Text,
									},
								},
							},
						},
					},
				}
				xmlRow.Cells = append(xmlRow.Cells, xmlCell)
			}

			xmlTable.Rows = append(xmlTable.Rows, xmlRow)
		}

		doc.Body.Tables = append(doc.Body.Tables, xmlTable)
	}

	// Marshal to XML
	var buf bytes.Buffer
	buf.WriteString(xml.Header)
	encoder := xml.NewEncoder(&buf)
	encoder.Indent("", "  ")
	if err := encoder.Encode(doc); err != nil {
		return nil, fmt.Errorf("failed to encode document XML: %w", err)
	}

	return buf.Bytes(), nil
}

// XML structures for document generation
type DocumentXML struct {
	XMLName xml.Name `xml:"w:document"`
	XMLNS   string   `xml:"xmlns:w,attr"`
	Body    DocumentBody
}

type DocumentBody struct {
	XMLName    xml.Name        `xml:"w:body"`
	Paragraphs []ParagraphXML  `xml:"w:p"`
	Tables     []TableXML      `xml:"w:tbl"`
}

type ParagraphXML struct {
	XMLName    xml.Name                `xml:"w:p"`
	Properties *ParagraphPropertiesXML `xml:"w:pPr,omitempty"`
	Runs       []RunXML               `xml:"w:r"`
}

type ParagraphPropertiesXML struct {
	XMLName xml.Name  `xml:"w:pPr"`
	Style   *StyleXML `xml:"w:pStyle,omitempty"`
}

type StyleXML struct {
	XMLName xml.Name `xml:"w:pStyle"`
	Val     string   `xml:"w:val,attr"`
}

type RunXML struct {
	XMLName    xml.Name         `xml:"w:r"`
	Properties *RunPropertiesXML `xml:"w:rPr,omitempty"`
	Text       *TextXML         `xml:"w:t,omitempty"`
}

type RunPropertiesXML struct {
	XMLName   xml.Name   `xml:"w:rPr"`
	Bold      *BoldXML   `xml:"w:b,omitempty"`
	Italic    *ItalicXML `xml:"w:i,omitempty"`
	Underline *UnderlineXML `xml:"w:u,omitempty"`
	Size      *SizeXML   `xml:"w:sz,omitempty"`
	Font      *FontXML   `xml:"w:rFonts,omitempty"`
}

type TextXML struct {
	XMLName xml.Name `xml:"w:t"`
	Content string   `xml:",chardata"`
}

type BoldXML struct {
	XMLName xml.Name `xml:"w:b"`
	Val     string   `xml:"w:val,attr"`
}

type ItalicXML struct {
	XMLName xml.Name `xml:"w:i"`
	Val     string   `xml:"w:val,attr"`
}

type UnderlineXML struct {
	XMLName xml.Name `xml:"w:u"`
	Val     string   `xml:"w:val,attr"`
}

type SizeXML struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr"`
}

type FontXML struct {
	XMLName xml.Name `xml:"w:rFonts"`
	Ascii   string   `xml:"w:ascii,attr"`
	HAnsi   string   `xml:"w:hAnsi,attr"`
}

type TableXML struct {
	XMLName xml.Name      `xml:"w:tbl"`
	Rows    []TableRowXML `xml:"w:tr"`
}

type TableRowXML struct {
	XMLName xml.Name        `xml:"w:tr"`
	Cells   []TableCellXML  `xml:"w:tc"`
}

type TableCellXML struct {
	XMLName xml.Name      `xml:"w:tc"`
	Content []interface{} `xml:",any"`
} 