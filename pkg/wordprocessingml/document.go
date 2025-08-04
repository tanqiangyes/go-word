// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/opc"
	"github.com/tanqiangyes/go-word/pkg/parser"
	"github.com/tanqiangyes/go-word/pkg/types"
)

// Document represents a Word document
type Document struct {
	container *opc.Container
	mainPart  *MainDocumentPart
	parts     map[string]*opc.Part
	documentParts *DocumentParts
}

// MainDocumentPart 使用 document_parts.go 中的定义

// Paragraph represents a paragraph in the document
type Paragraph = types.Paragraph

// Run represents a text run with specific formatting
type Run = types.Run

// Table represents a table in the document
type Table = types.Table

// TableRow represents a row in a table
type TableRow = types.TableRow

// TableCell represents a cell in a table
type TableCell = types.TableCell

// Open opens a Word document from a file
func Open(filename string) (*Document, error) {
	container, err := opc.Open(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open document: %w", err)
	}
	
	doc := &Document{
		container: container,
		parts:     make(map[string]*opc.Part),
		documentParts: NewDocumentParts(),
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

// GetContainer returns the underlying OPC container
func (d *Document) GetContainer() *opc.Container {
	return d.container
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

// GetDocumentParts returns the document parts
func (d *Document) GetDocumentParts() *DocumentParts {
	return d.documentParts
}

// GetPartsSummary returns a summary of all document parts
func (d *Document) GetPartsSummary() string {
	if d.documentParts == nil {
		return "文档部分未加载"
	}
	return d.documentParts.GetPartsSummary()
}

// loadMainDocumentPart loads the main document part
func (d *Document) loadMainDocumentPart() error {
	// Get the main document part
	mainPart, err := d.container.GetPart("word/document.xml")
	if err != nil {
		return fmt.Errorf("failed to get main document part: %w", err)
	}
	
	d.mainPart = &MainDocumentPart{
		ID: "main",
	}
	
	// Parse the document content
	content, err := parseDocumentContent(mainPart.Content)
	if err != nil {
		return fmt.Errorf("failed to parse document content: %w", err)
	}
	
	d.mainPart.Content = content
	
	// 设置主文档部分到文档部分
	d.documentParts.MainDocumentPart = d.mainPart
	return nil
}

// parseDocumentContent parses the XML content of the document
func parseDocumentContent(content []byte) (*types.DocumentContent, error) {
	wordParser := &parser.WordMLParser{}
	
	// Parse the XML content
	docXML, err := wordParser.ParseWordDocument(content)
	if err != nil {
		return nil, fmt.Errorf("failed to parse document XML: %w", err)
	}
	
	// Extract content
	text := wordParser.ExtractText(docXML)
	paragraphs := wordParser.ExtractParagraphs(docXML)
	tables := wordParser.ExtractTables(docXML)
	
	return &types.DocumentContent{
		Paragraphs: paragraphs,
		Tables:     tables,
		Text:       text,
	}, nil
}

// GetMainPart returns the main document part
func (d *Document) GetMainPart() *MainDocumentPart {
	return d.mainPart
}

// SetMainPart sets the main document part
func (d *Document) SetMainPart(mainPart *MainDocumentPart) {
	d.mainPart = mainPart
} 