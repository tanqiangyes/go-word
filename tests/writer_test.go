package tests

import (
	"os"
	"testing"

	"github.com/go-word/pkg/writer"
	"github.com/go-word/pkg/types"
)

func TestDocumentWriterCreateNewDocument(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}
}

func TestDocumentWriterAddParagraph(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}

	err = docWriter.AddParagraph("Test paragraph", "Normal")
	if err != nil {
		t.Fatalf("Failed to add paragraph: %v", err)
	}
}

func TestDocumentWriterAddTable(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}

	tableData := [][]string{
		{"Header1", "Header2"},
		{"Cell1", "Cell2"},
	}

	err = docWriter.AddTable(tableData)
	if err != nil {
		t.Fatalf("Failed to add table: %v", err)
	}
}

func TestDocumentWriterAddFormattedParagraph(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}

	formattedRuns := []types.Run{
		{
			Text:     "Bold text",
			Bold:     true,
			FontSize: 16,
		},
		{
			Text:     "Italic text",
			Italic:   true,
			FontSize: 14,
		},
	}

	err = docWriter.AddFormattedParagraph("Formatted paragraph", "Normal", formattedRuns)
	if err != nil {
		t.Fatalf("Failed to add formatted paragraph: %v", err)
	}
}

func TestDocumentWriterReplaceText(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}

	err = docWriter.AddParagraph("Original text", "Normal")
	if err != nil {
		t.Fatalf("Failed to add paragraph: %v", err)
	}

	err = docWriter.ReplaceText("Original", "Modified")
	if err != nil {
		t.Fatalf("Failed to replace text: %v", err)
	}
}

func TestDocumentWriterSave(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}

	err = docWriter.AddParagraph("Test content", "Normal")
	if err != nil {
		t.Fatalf("Failed to add paragraph: %v", err)
	}

	// Save to temporary file
	tempFile := "test_output.docx"
	defer os.Remove(tempFile)

	err = docWriter.Save(tempFile)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Check if file was created
	if _, err := os.Stat(tempFile); os.IsNotExist(err) {
		t.Error("Document file was not created")
	}
}

func TestDocumentWriterSetParagraphStyle(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}

	err = docWriter.AddParagraph("Test paragraph", "Normal")
	if err != nil {
		t.Fatalf("Failed to add paragraph: %v", err)
	}

	err = docWriter.SetParagraphStyle(0, "Heading1")
	if err != nil {
		t.Fatalf("Failed to set paragraph style: %v", err)
	}
}

func TestDocumentWriterSetRunFormatting(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create new document: %v", err)
	}

	err = docWriter.AddParagraph("Test paragraph", "Normal")
	if err != nil {
		t.Fatalf("Failed to add paragraph: %v", err)
	}

	formatting := types.Run{
		Text:     "Formatted text",
		Bold:     true,
		Italic:   true,
		FontSize: 18,
		FontName: "Arial",
	}

	err = docWriter.SetRunFormatting(0, 0, formatting)
	if err != nil {
		t.Fatalf("Failed to set run formatting: %v", err)
	}
}

func TestDocumentWriterErrorHandling(t *testing.T) {
	docWriter := writer.NewDocumentWriter()

	// Test operations without initializing document
	err := docWriter.AddParagraph("Test", "Normal")
	if err == nil {
		t.Error("Expected error when adding paragraph to uninitialized document")
	}

	err = docWriter.AddTable([][]string{{"test"}})
	if err == nil {
		t.Error("Expected error when adding table to uninitialized document")
	}

	err = docWriter.ReplaceText("old", "new")
	if err == nil {
		t.Error("Expected error when replacing text in uninitialized document")
	}

	err = docWriter.SetParagraphStyle(0, "Normal")
	if err == nil {
		t.Error("Expected error when setting style in uninitialized document")
	}

	err = docWriter.Save("test.docx")
	if err == nil {
		t.Error("Expected error when saving uninitialized document")
	}
} 