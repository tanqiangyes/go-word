// Package types provides shared type definitions for the go-word library
package types

import "encoding/xml"

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
	Color    string
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

// 通用Word格式属性类型
type Bold struct {
	XMLName xml.Name `xml:"w:b"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

type Italic struct {
	XMLName xml.Name `xml:"w:i"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

type Size struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr"`
}

type Font struct {
	XMLName xml.Name `xml:"w:rFonts"`
	Ascii   string   `xml:"w:ascii,attr,omitempty"`
	HAnsi   string   `xml:"w:hAnsi,attr,omitempty"`
}

type Underline struct {
	XMLName xml.Name `xml:"w:u"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr,omitempty"`
} 