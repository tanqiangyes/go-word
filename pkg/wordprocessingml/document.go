// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"strings"
	
	"github.com/go-word/pkg/opc"
)

// Document represents a Word document
type Document struct {
	container *opc.Container
	mainPart  *MainDocumentPart
	parts     map[string]*opc.Part
}

// MainDocumentPart represents the main document part
type MainDocumentPart struct {
	Part     *opc.Part
	Content  *DocumentContent
}

// DocumentContent represents the content of the document
type DocumentContent struct {
	Paragraphs []Paragraph
	Tables     []Table
	Text       string
}

// Paragraph represents a paragraph in the document
type Paragraph struct {
	Text   string
	Style  string
	Runs   []Run
}

// Run represents a text run with specific formatting
type Run struct {
	Text     string
	Bold     bool
	Italic   bool
	Underline bool
	FontSize int
	FontName string
}

// Table represents a table in the document
type Table struct {
	Rows    []TableRow
	Columns int
}

// TableRow represents a row in a table
type TableRow struct {
	Cells []TableCell
}

// TableCell represents a cell in a table
type TableCell struct {
	Text string
}

// Open opens a Word document from a file
func Open(filename string) (*Document, error) {
	container, err := opc.Open(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open document: %w", err)
	}
	
	doc := &Document{
		container: container,
		parts:     make(map[string]*opc.Part),
	}
	
	// Load main document part
	if err := doc.loadMainDocumentPart(); err != nil {
		container.Close()
		return nil, fmt.Errorf("failed to load main document part: %w", err)
	}
	
	return doc, nil
}

// Close closes the document and releases resources
func (d *Document) Close() error {
	if d.container != nil {
		return d.container.Close()
	}
	return nil
}

// GetText returns the plain text content of the document
func (d *Document) GetText() (string, error) {
	if d.mainPart == nil || d.mainPart.Content == nil {
		return "", fmt.Errorf("document content not loaded")
	}
	
	var text strings.Builder
	for _, paragraph := range d.mainPart.Content.Paragraphs {
		text.WriteString(paragraph.Text)
		text.WriteString("\n")
	}
	
	return text.String(), nil
}

// GetParagraphs returns all paragraphs in the document
func (d *Document) GetParagraphs() ([]Paragraph, error) {
	if d.mainPart == nil || d.mainPart.Content == nil {
		return nil, fmt.Errorf("document content not loaded")
	}
	
	return d.mainPart.Content.Paragraphs, nil
}

// GetTables returns all tables in the document
func (d *Document) GetTables() ([]Table, error) {
	if d.mainPart == nil || d.mainPart.Content == nil {
		return nil, fmt.Errorf("document content not loaded")
	}
	
	return d.mainPart.Content.Tables, nil
}

// loadMainDocumentPart loads the main document part
func (d *Document) loadMainDocumentPart() error {
	// Get the main document part
	mainPart, err := d.container.GetPart("word/document.xml")
	if err != nil {
		return fmt.Errorf("failed to get main document part: %w", err)
	}
	
	d.mainPart = &MainDocumentPart{
		Part: mainPart,
	}
	
	// Parse the document content
	content, err := parseDocumentContent(mainPart.Content)
	if err != nil {
		return fmt.Errorf("failed to parse document content: %w", err)
	}
	
	d.mainPart.Content = content
	return nil
}

// parseDocumentContent parses the XML content of the document
func parseDocumentContent(content []byte) (*DocumentContent, error) {
	// TODO: Implement XML parsing for document content
	// For now, return a basic structure
	return &DocumentContent{
		Paragraphs: []Paragraph{},
		Tables:     []Table{},
		Text:       "",
	}, nil
} 