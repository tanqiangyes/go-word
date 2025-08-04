// Package types provides shared type definitions for the go-word library
package types

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

// DocumentContent represents the content of the document
type DocumentContent struct {
	Paragraphs []Paragraph
	Tables     []Table
	Text       string
} 