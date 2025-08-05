package types

import (
	"testing"
)

func TestParagraph(t *testing.T) {
	// Test creating a paragraph
	paragraph := Paragraph{
		Text:  "Test paragraph text",
		Style: "Normal",
		Runs: []Run{
			{
				Text:     "Bold text",
				Bold:     true,
				FontSize: 12,
				FontName: "Arial",
			},
			{
				Text:     "Italic text",
				Italic:   true,
				FontSize: 10,
				FontName: "Times New Roman",
			},
		},
	}

	// Test basic properties
	if paragraph.Text != "Test paragraph text" {
		t.Errorf("Expected text 'Test paragraph text', got '%s'", paragraph.Text)
	}

	if paragraph.Style != "Normal" {
		t.Errorf("Expected style 'Normal', got '%s'", paragraph.Style)
	}

	if len(paragraph.Runs) != 2 {
		t.Errorf("Expected 2 runs, got %d", len(paragraph.Runs))
	}

	// Test first run
	firstRun := paragraph.Runs[0]
	if firstRun.Text != "Bold text" {
		t.Errorf("Expected first run text 'Bold text', got '%s'", firstRun.Text)
	}

	if !firstRun.Bold {
		t.Error("Expected first run to be bold")
	}

	if firstRun.FontSize != 12 {
		t.Errorf("Expected font size 12, got %d", firstRun.FontSize)
	}

	// Test second run
	secondRun := paragraph.Runs[1]
	if secondRun.Text != "Italic text" {
		t.Errorf("Expected second run text 'Italic text', got '%s'", secondRun.Text)
	}

	if !secondRun.Italic {
		t.Error("Expected second run to be italic")
	}

	if secondRun.FontSize != 10 {
		t.Errorf("Expected font size 10, got %d", secondRun.FontSize)
	}
}

func TestRun(t *testing.T) {
	// Test creating a run with all properties
	run := Run{
		Text:      "Test run text",
		Bold:      true,
		Italic:    true,
		Underline: true,
		FontSize:  14,
		FontName:  "Calibri",
		Color:     "FF0000",
	}

	// Test all properties
	if run.Text != "Test run text" {
		t.Errorf("Expected text 'Test run text', got '%s'", run.Text)
	}

	if !run.Bold {
		t.Error("Expected run to be bold")
	}

	if !run.Italic {
		t.Error("Expected run to be italic")
	}

	if !run.Underline {
		t.Error("Expected run to be underlined")
	}

	if run.FontSize != 14 {
		t.Errorf("Expected font size 14, got %d", run.FontSize)
	}

	if run.FontName != "Calibri" {
		t.Errorf("Expected font name 'Calibri', got '%s'", run.FontName)
	}

	if run.Color != "FF0000" {
		t.Errorf("Expected color 'FF0000', got '%s'", run.Color)
	}

	// Test default values
	defaultRun := Run{}
	if defaultRun.Text != "" {
		t.Errorf("Expected empty text for default run, got '%s'", defaultRun.Text)
	}

	if defaultRun.Bold {
		t.Error("Expected default run to not be bold")
	}

	if defaultRun.Italic {
		t.Error("Expected default run to not be italic")
	}

	if defaultRun.Underline {
		t.Error("Expected default run to not be underlined")
	}

	if defaultRun.FontSize != 0 {
		t.Errorf("Expected default font size 0, got %d", defaultRun.FontSize)
	}
}

func TestTable(t *testing.T) {
	// Test creating a table
	table := Table{
		Columns: 3,
		Rows: []TableRow{
			{
				Cells: []TableCell{
					{Text: "Header 1"},
					{Text: "Header 2"},
					{Text: "Header 3"},
				},
			},
			{
				Cells: []TableCell{
					{Text: "Cell 1"},
					{Text: "Cell 2"},
					{Text: "Cell 3"},
				},
			},
		},
	}

	// Test basic properties
	if table.Columns != 3 {
		t.Errorf("Expected 3 columns, got %d", table.Columns)
	}

	if len(table.Rows) != 2 {
		t.Errorf("Expected 2 rows, got %d", len(table.Rows))
	}

	// Test first row (header)
	headerRow := table.Rows[0]
	if len(headerRow.Cells) != 3 {
		t.Errorf("Expected 3 cells in header row, got %d", len(headerRow.Cells))
	}

	if headerRow.Cells[0].Text != "Header 1" {
		t.Errorf("Expected first header cell 'Header 1', got '%s'", headerRow.Cells[0].Text)
	}

	// Test second row (data)
	dataRow := table.Rows[1]
	if len(dataRow.Cells) != 3 {
		t.Errorf("Expected 3 cells in data row, got %d", len(dataRow.Cells))
	}

	if dataRow.Cells[0].Text != "Cell 1" {
		t.Errorf("Expected first data cell 'Cell 1', got '%s'", dataRow.Cells[0].Text)
	}
}

func TestTableRow(t *testing.T) {
	// Test creating a table row
	row := TableRow{
		Cells: []TableCell{
			{Text: "First cell"},
			{Text: "Second cell"},
			{Text: "Third cell"},
		},
	}

	// Test cell count
	if len(row.Cells) != 3 {
		t.Errorf("Expected 3 cells, got %d", len(row.Cells))
	}

	// Test cell contents
	expectedTexts := []string{"First cell", "Second cell", "Third cell"}
	for i, expected := range expectedTexts {
		if row.Cells[i].Text != expected {
			t.Errorf("Expected cell %d text '%s', got '%s'", i+1, expected, row.Cells[i].Text)
		}
	}

	// Test empty row
	emptyRow := TableRow{}
	if len(emptyRow.Cells) != 0 {
		t.Errorf("Expected empty row to have 0 cells, got %d", len(emptyRow.Cells))
	}
}

func TestTableCell(t *testing.T) {
	// Test creating a table cell
	cell := TableCell{
		Text: "Test cell content",
	}

	// Test cell text
	if cell.Text != "Test cell content" {
		t.Errorf("Expected cell text 'Test cell content', got '%s'", cell.Text)
	}

	// Test empty cell
	emptyCell := TableCell{}
	if emptyCell.Text != "" {
		t.Errorf("Expected empty cell to have empty text, got '%s'", emptyCell.Text)
	}
}

func TestDocumentContent(t *testing.T) {
	// Test creating document content
	content := DocumentContent{
		Text: "Document text content",
		Paragraphs: []Paragraph{
			{
				Text:  "First paragraph",
				Style: "Normal",
			},
			{
				Text:  "Second paragraph",
				Style: "Heading1",
			},
		},
		Tables: []Table{
			{
				Columns: 2,
				Rows: []TableRow{
					{
						Cells: []TableCell{
							{Text: "Header 1"},
							{Text: "Header 2"},
						},
					},
				},
			},
		},
	}

	// Test text content
	if content.Text != "Document text content" {
		t.Errorf("Expected text 'Document text content', got '%s'", content.Text)
	}

	// Test paragraphs
	if len(content.Paragraphs) != 2 {
		t.Errorf("Expected 2 paragraphs, got %d", len(content.Paragraphs))
	}

	if content.Paragraphs[0].Text != "First paragraph" {
		t.Errorf("Expected first paragraph 'First paragraph', got '%s'", content.Paragraphs[0].Text)
	}

	if content.Paragraphs[1].Style != "Heading1" {
		t.Errorf("Expected second paragraph style 'Heading1', got '%s'", content.Paragraphs[1].Style)
	}

	// Test tables
	if len(content.Tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(content.Tables))
	}

	if content.Tables[0].Columns != 2 {
		t.Errorf("Expected table to have 2 columns, got %d", content.Tables[0].Columns)
	}

	// Test empty content
	emptyContent := DocumentContent{}
	if emptyContent.Text != "" {
		t.Errorf("Expected empty content to have empty text, got '%s'", emptyContent.Text)
	}

	if len(emptyContent.Paragraphs) != 0 {
		t.Errorf("Expected empty content to have 0 paragraphs, got %d", len(emptyContent.Paragraphs))
	}

	if len(emptyContent.Tables) != 0 {
		t.Errorf("Expected empty content to have 0 tables, got %d", len(emptyContent.Tables))
	}
}

func TestWordFormatTypes(t *testing.T) {
	// Test Bold type
	bold := Bold{Val: "true"}
	if bold.Val != "true" {
		t.Errorf("Expected Bold.Val 'true', got '%s'", bold.Val)
	}

	// Test Italic type
	italic := Italic{Val: "false"}
	if italic.Val != "false" {
		t.Errorf("Expected Italic.Val 'false', got '%s'", italic.Val)
	}

	// Test Size type
	size := Size{Val: "24"}
	if size.Val != "24" {
		t.Errorf("Expected Size.Val '24', got '%s'", size.Val)
	}

	// Test Font type
	font := Font{
		Ascii: "Arial",
		HAnsi: "Arial",
	}
	if font.Ascii != "Arial" {
		t.Errorf("Expected Font.Ascii 'Arial', got '%s'", font.Ascii)
	}

	if font.HAnsi != "Arial" {
		t.Errorf("Expected Font.HAnsi 'Arial', got '%s'", font.HAnsi)
	}

	// Test Underline type
	underline := Underline{Val: "single"}
	if underline.Val != "single" {
		t.Errorf("Expected Underline.Val 'single', got '%s'", underline.Val)
	}

	// Test Color type
	color := Color{Val: "FF0000"}
	if color.Val != "FF0000" {
		t.Errorf("Expected Color.Val 'FF0000', got '%s'", color.Val)
	}
}

func TestWordFormatTypesWithEmptyValues(t *testing.T) {
	// Test types with empty values
	bold := Bold{}
	if bold.Val != "" {
		t.Errorf("Expected empty Bold.Val, got '%s'", bold.Val)
	}

	italic := Italic{}
	if italic.Val != "" {
		t.Errorf("Expected empty Italic.Val, got '%s'", italic.Val)
	}

	size := Size{}
	if size.Val != "" {
		t.Errorf("Expected empty Size.Val, got '%s'", size.Val)
	}

	font := Font{}
	if font.Ascii != "" {
		t.Errorf("Expected empty Font.Ascii, got '%s'", font.Ascii)
	}

	if font.HAnsi != "" {
		t.Errorf("Expected empty Font.HAnsi, got '%s'", font.HAnsi)
	}

	underline := Underline{}
	if underline.Val != "" {
		t.Errorf("Expected empty Underline.Val, got '%s'", underline.Val)
	}

	color := Color{}
	if color.Val != "" {
		t.Errorf("Expected empty Color.Val, got '%s'", color.Val)
	}
} 
