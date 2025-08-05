package types

import (
	"testing"
)

func TestParagraphCreation(t *testing.T) {
	paragraph := Paragraph{
		Text:  "Test paragraph",
		Style: "Normal",
		Runs: []Run{
			{Text: "Test run"},
		},
	}
	
	if paragraph.Text != "Test paragraph" {
		t.Errorf("Expected text 'Test paragraph', got '%s'", paragraph.Text)
	}
	
	if paragraph.Style != "Normal" {
		t.Errorf("Expected style 'Normal', got '%s'", paragraph.Style)
	}
	
	if len(paragraph.Runs) != 1 {
		t.Errorf("Expected 1 run, got %d", len(paragraph.Runs))
	}
}

func TestRunCreation(t *testing.T) {
	run := Run{
		Text:      "Test text",
		Bold:      true,
		Italic:    false,
		Underline: true,
		FontSize:  12,
		FontName:  "Arial",
		Color:     "000000",
	}
	
	if run.Text != "Test text" {
		t.Errorf("Expected text 'Test text', got '%s'", run.Text)
	}
	
	if !run.Bold {
		t.Error("Expected Bold to be true")
	}
	
	if run.Italic {
		t.Error("Expected Italic to be false")
	}
	
	if !run.Underline {
		t.Error("Expected Underline to be true")
	}
	
	if run.FontSize != 12 {
		t.Errorf("Expected font size 12, got %d", run.FontSize)
	}
	
	if run.FontName != "Arial" {
		t.Errorf("Expected font name 'Arial', got '%s'", run.FontName)
	}
	
	if run.Color != "000000" {
		t.Errorf("Expected color '000000', got '%s'", run.Color)
	}
}

func TestTableCreation(t *testing.T) {
	table := Table{
		Rows: []TableRow{
			{
				Cells: []TableCell{
					{Text: "Cell 1"},
					{Text: "Cell 2"},
				},
			},
			{
				Cells: []TableCell{
					{Text: "Cell 3"},
					{Text: "Cell 4"},
				},
			},
		},
		Columns: 2,
	}
	
	if len(table.Rows) != 2 {
		t.Errorf("Expected 2 rows, got %d", len(table.Rows))
	}
	
	if table.Columns != 2 {
		t.Errorf("Expected 2 columns, got %d", table.Columns)
	}
	
	if len(table.Rows[0].Cells) != 2 {
		t.Errorf("Expected 2 cells in first row, got %d", len(table.Rows[0].Cells))
	}
}

func TestTableRowCreation(t *testing.T) {
	row := TableRow{
		Cells: []TableCell{
			{Text: "Header 1"},
			{Text: "Header 2"},
			{Text: "Header 3"},
		},
	}
	
	if len(row.Cells) != 3 {
		t.Errorf("Expected 3 cells, got %d", len(row.Cells))
	}
	
	if row.Cells[0].Text != "Header 1" {
		t.Errorf("Expected first cell 'Header 1', got '%s'", row.Cells[0].Text)
	}
}

func TestTableCellCreation(t *testing.T) {
	cell := TableCell{
		Text: "Cell content",
	}
	
	if cell.Text != "Cell content" {
		t.Errorf("Expected text 'Cell content', got '%s'", cell.Text)
	}
}

func TestDocumentContentCreation(t *testing.T) {
	content := &DocumentContent{
		Paragraphs: []Paragraph{
			{Text: "First paragraph"},
			{Text: "Second paragraph"},
		},
		Tables: []Table{
			{Rows: []TableRow{}},
		},
		Text: "Document text",
	}
	
	if len(content.Paragraphs) != 2 {
		t.Errorf("Expected 2 paragraphs, got %d", len(content.Paragraphs))
	}
	
	if len(content.Tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(content.Tables))
	}
	
	if content.Text != "Document text" {
		t.Errorf("Expected text 'Document text', got '%s'", content.Text)
	}
}

func TestBoldXML(t *testing.T) {
	bold := Bold{
		Val: "true",
	}
	
	if bold.Val != "true" {
		t.Errorf("Expected Val 'true', got '%s'", bold.Val)
	}
}

func TestItalicXML(t *testing.T) {
	italic := Italic{
		Val: "true",
	}
	
	if italic.Val != "true" {
		t.Errorf("Expected Val 'true', got '%s'", italic.Val)
	}
}

func TestSizeXML(t *testing.T) {
	size := Size{
		Val: "24",
	}
	
	if size.Val != "24" {
		t.Errorf("Expected Val '24', got '%s'", size.Val)
	}
}

func TestFontXML(t *testing.T) {
	font := Font{
		Ascii: "Arial",
		HAnsi: "Arial",
	}
	
	if font.Ascii != "Arial" {
		t.Errorf("Expected Ascii 'Arial', got '%s'", font.Ascii)
	}
	
	if font.HAnsi != "Arial" {
		t.Errorf("Expected HAnsi 'Arial', got '%s'", font.HAnsi)
	}
}

func TestUnderlineXML(t *testing.T) {
	underline := Underline{
		Val: "single",
	}
	
	if underline.Val != "single" {
		t.Errorf("Expected Val 'single', got '%s'", underline.Val)
	}
}

func TestColorXML(t *testing.T) {
	color := Color{
		Val: "FF0000",
	}
	
	if color.Val != "FF0000" {
		t.Errorf("Expected Val 'FF0000', got '%s'", color.Val)
	}
}

func TestEmptyDocumentContent(t *testing.T) {
	content := &DocumentContent{}
	
	if len(content.Paragraphs) != 0 {
		t.Errorf("Expected 0 paragraphs, got %d", len(content.Paragraphs))
	}
	
	if len(content.Tables) != 0 {
		t.Errorf("Expected 0 tables, got %d", len(content.Tables))
	}
	
	if content.Text != "" {
		t.Errorf("Expected empty text, got '%s'", content.Text)
	}
}

func TestRunWithAllProperties(t *testing.T) {
	run := Run{
		Text:      "Formatted text",
		Bold:      true,
		Italic:    true,
		Underline: true,
		FontSize:  16,
		FontName:  "Times New Roman",
		Color:     "0000FF",
	}
	
	if !run.Bold {
		t.Error("Expected Bold to be true")
	}
	
	if !run.Italic {
		t.Error("Expected Italic to be true")
	}
	
	if !run.Underline {
		t.Error("Expected Underline to be true")
	}
	
	if run.FontSize != 16 {
		t.Errorf("Expected font size 16, got %d", run.FontSize)
	}
	
	if run.FontName != "Times New Roman" {
		t.Errorf("Expected font name 'Times New Roman', got '%s'", run.FontName)
	}
	
	if run.Color != "0000FF" {
		t.Errorf("Expected color '0000FF', got '%s'", run.Color)
	}
} 