// Package wordprocessingml provides WordprocessingML document processing functionality.
// This package implements the core functionality for reading, parsing, and manipulating
// Word documents (.docx files) according to the Office Open XML specification.
//
// The package provides high-level APIs for:
//   - Opening and closing Word documents
//   - Extracting text content and formatting
//   - Accessing paragraphs, tables, and other document elements
//   - Advanced formatting and document manipulation
//
// Example usage:
//
//	doc, err := wordprocessingml.Open("document.docx")
//	if err != nil {
//		log.Fatal(err)
//	}
//	defer doc.Close()
//
//	text, err := doc.GetText()
//	if err != nil {
//		log.Fatal(err)
//	}
//	fmt.Println("Document content:", text)
package wordprocessingml

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/opc"
	"github.com/tanqiangyes/go-word/pkg/parser"
	"github.com/tanqiangyes/go-word/pkg/types"
)

// Document represents a Word document and provides methods for accessing
// and manipulating its content. A Document instance should be closed
// when no longer needed to free up system resources.
//
// The Document struct contains:
//   - container: The underlying OPC container that holds the document files
//   - mainPart: The main document part containing the actual content
//   - parts: A map of all document parts (headers, footers, styles, etc.)
//   - documentParts: A manager for document parts and their relationships
type Document struct {
	container     *opc.Container
	mainPart      *MainDocumentPart
	parts         map[string]*opc.Part
	documentParts *DocumentParts
}

// MainDocumentPart represents the main content of a Word document.
// It contains the actual text, paragraphs, tables, and other content
// that appears in the document body.

// Paragraph represents a paragraph in the document.
// A paragraph contains text runs with specific formatting applied.
type Paragraph = types.Paragraph

// Run represents a text run with specific formatting.
// A run is a continuous segment of text that shares the same formatting
// properties (bold, italic, font size, etc.).
type Run = types.Run

// Table represents a table in the document.
// Tables are organized in rows and columns, with each cell containing
// text content and optional formatting.
type Table = types.Table

// TableRow represents a row in a table.
// Each row contains multiple cells arranged horizontally.
type TableRow = types.TableRow

// TableCell represents a cell in a table.
// A cell can contain text content and may be merged with other cells.
type TableCell = types.TableCell

// Open opens a Word document from a file and returns a Document instance.
// The returned document must be closed when no longer needed to free up
// system resources.
//
// Parameters:
//   - filename: The path to the .docx file to open
//
// Returns:
//   - *Document: A document instance that can be used to access content
//   - error: An error if the file cannot be opened or is not a valid Word document
//
// Example:
//
//	doc, err := wordprocessingml.Open("document.docx")
//	if err != nil {
//		log.Fatal("Failed to open document:", err)
//	}
//	defer doc.Close()
//
// Note: This function will attempt to parse the document structure and may
// return an error if the document is corrupted or uses unsupported features.
func Open(filename string) (*Document, error) {
	// Open the OPC container (ZIP file) that contains the Word document
	container, err := opc.Open(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open document: %w", err)
	}
	
	// Create a new document instance
	doc := &Document{
		container:     container,
		parts:         make(map[string]*opc.Part),
		documentParts: NewDocumentParts(),
	}
	
	// Load the main document part (word/document.xml)
	if err := doc.loadMainDocumentPart(); err != nil {
		container.Close()
		return nil, fmt.Errorf("failed to load main document part: %w", err)
	}
	
	return doc, nil
}

// Close closes the document and releases all associated resources.
// This method should be called when the document is no longer needed
// to prevent memory leaks and file handle exhaustion.
//
// Returns:
//   - error: An error if the document cannot be closed properly
//
// Example:
//
//	doc, err := wordprocessingml.Open("document.docx")
//	if err != nil {
//		log.Fatal(err)
//	}
//	defer doc.Close() // Ensure resources are freed
func (d *Document) Close() error {
	if d.container != nil {
		return d.container.Close()
	}
	return nil
}

// GetContainer returns the underlying OPC container.
// This method provides access to the low-level container for advanced
// operations that require direct access to document parts.
//
// Returns:
//   - *opc.Container: The underlying OPC container
//
// Note: This method is intended for advanced users who need to perform
// custom operations on the document structure.
func (d *Document) GetContainer() *opc.Container {
	return d.container
}

// GetText returns the plain text content of the document.
// This method extracts all text from the document, including text from
// paragraphs and tables, but excludes formatting information.
//
// Returns:
//   - string: The plain text content of the document
//   - error: An error if the text cannot be extracted
//
// Example:
//
//	text, err := doc.GetText()
//	if err != nil {
//		log.Fatal("Failed to get text:", err)
//	}
//	fmt.Printf("Document contains %d characters\n", len(text))
//
// Note: The returned text includes newline characters between paragraphs
// and may not preserve the exact formatting of the original document.
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

// GetParagraphs returns all paragraphs in the document.
// Each paragraph contains text runs with formatting information.
//
// Returns:
//   - []Paragraph: A slice of all paragraphs in the document
//   - error: An error if the paragraphs cannot be retrieved
//
// Example:
//
//	paragraphs, err := doc.GetParagraphs()
//	if err != nil {
//		log.Fatal("Failed to get paragraphs:", err)
//	}
//	
//	for i, paragraph := range paragraphs {
//		fmt.Printf("Paragraph %d: %s\n", i+1, paragraph.Text)
//		for j, run := range paragraph.Runs {
//			fmt.Printf("  Run %d: '%s' (Bold: %v, Italic: %v)\n",
//				j+1, run.Text, run.Bold, run.Italic)
//		}
//	}
func (d *Document) GetParagraphs() ([]Paragraph, error) {
	if d.mainPart == nil || d.mainPart.Content == nil {
		return nil, fmt.Errorf("document content not loaded")
	}
	
	return d.mainPart.Content.Paragraphs, nil
}

// GetTables returns all tables in the document.
// Each table contains rows and columns with text content.
//
// Returns:
//   - []Table: A slice of all tables in the document
//   - error: An error if the tables cannot be retrieved
//
// Example:
//
//	tables, err := doc.GetTables()
//	if err != nil {
//		log.Fatal("Failed to get tables:", err)
//	}
//	
//	for i, table := range tables {
//		fmt.Printf("Table %d: %d rows x %d columns\n",
//			i+1, len(table.Rows), table.Columns)
//		
//		for rowIdx, row := range table.Rows {
//			for colIdx, cell := range row.Cells {
//				fmt.Printf("  Cell [%d,%d]: %s\n", rowIdx, colIdx, cell.Text)
//			}
//		}
//	}
func (d *Document) GetTables() ([]Table, error) {
	if d.mainPart == nil || d.mainPart.Content == nil {
		return nil, fmt.Errorf("document content not loaded")
	}
	
	return d.mainPart.Content.Tables, nil
}

// GetDocumentParts returns the document parts manager.
// This provides access to document parts like headers, footers,
// styles, and other document components.
//
// Returns:
//   - *DocumentParts: The document parts manager
func (d *Document) GetDocumentParts() *DocumentParts {
	return d.documentParts
}

// GetPartsSummary returns a summary of all document parts.
// This is useful for debugging and understanding the document structure.
//
// Returns:
//   - string: A formatted summary of all document parts
func (d *Document) GetPartsSummary() string {
	if d.container == nil {
		return "Document not loaded"
	}
	
	parts, err := d.container.ListParts()
	if err != nil {
		return fmt.Sprintf("Error listing parts: %v", err)
	}
	
	var summary strings.Builder
	summary.WriteString(fmt.Sprintf("Document contains %d parts:\n", len(parts)))
	
	for _, partName := range parts {
		part, err := d.container.GetPart(partName)
		if err != nil {
			summary.WriteString(fmt.Sprintf("  %s (error: %v)\n", partName, err))
		} else {
			summary.WriteString(fmt.Sprintf("  %s (%d bytes)\n", partName, len(part.Content)))
		}
	}
	
	return summary.String()
}

// loadMainDocumentPart loads the main document part from the container.
// This method is called internally by Open() and should not be called
// directly by users.
//
// Returns:
//   - error: An error if the main document part cannot be loaded
func (d *Document) loadMainDocumentPart() error {
	// Get the main document part from the container
	part, err := d.container.GetPart("word/document.xml")
	if err != nil {
		return fmt.Errorf("failed to get main document part: %w", err)
	}
	
	// Parse the document content
	content, err := parseDocumentContent(part.Content)
	if err != nil {
		return fmt.Errorf("failed to parse document content: %w", err)
	}
	
	// Create the main document part
	d.mainPart = &MainDocumentPart{
		Content: content,
	}
	
	return nil
}

// parseDocumentContent parses the XML content of the main document part.
// This function converts the WordprocessingML XML into structured data
// that can be easily accessed by the Document methods.
//
// Parameters:
//   - content: The XML data from the main document part
//
// Returns:
//   - *types.DocumentContent: The parsed document content
//   - error: An error if the XML cannot be parsed
func parseDocumentContent(content []byte) (*types.DocumentContent, error) {
	// Parse the WordprocessingML XML
	docContent, err := parser.ParseWordML(content)
	if err != nil {
		return nil, fmt.Errorf("failed to parse WordML: %w", err)
	}
	
	return docContent, nil
}

// GetMainPart returns the main document part.
// This method provides access to the internal document structure
// and is primarily used by advanced formatting features.
//
// Returns:
//   - *MainDocumentPart: The main document part
func (d *Document) GetMainPart() *MainDocumentPart {
	return d.mainPart
}

// SetMainPart sets the main document part.
// This method is used internally by advanced features and should
// not be called directly by users unless they understand the
// document structure.
//
// Parameters:
//   - mainPart: The main document part to set
func (d *Document) SetMainPart(mainPart *MainDocumentPart) {
	d.mainPart = mainPart
} 