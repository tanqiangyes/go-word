package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
	fmt.Println("Go-Word Package Import Test")
	fmt.Println("============================")

	// Test creating a new document
	fmt.Println("\n1. Creating a new document...")
	doc, err := word.New()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}
	defer doc.Close()

	fmt.Println("✓ Document created successfully")

	// Test accessing types from the types package
	fmt.Println("\n2. Testing type access...")

	// Create a sample paragraph
	paragraph := types.Paragraph{
		Text: "This is a test paragraph",
		Runs: []types.Run{
			{
				Text: "This is a test paragraph",
				Bold: true,
			},
		},
	}

	fmt.Printf("✓ Created paragraph with text: %s\n", paragraph.Text)
	fmt.Printf("✓ Paragraph has %d runs\n", len(paragraph.Runs))
	fmt.Printf("✓ First run is bold: %v\n", paragraph.Runs[0].Bold)

	// Test document methods
	fmt.Println("\n3. Testing document methods...")

	// Get document parts summary
	summary := doc.GetPartsSummary()
	fmt.Println("Document summary:", summary)

	// Get paragraphs (should be empty for new document)
	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		log.Fatal("Failed to get paragraphs:", err)
	}
	fmt.Printf("✓ Found %d paragraphs\n", len(paragraphs))

	// Get tables (should be empty for new document)
	tables, err := doc.GetTables()
	if err != nil {
		log.Fatal("Failed to get tables:", err)
	}
	fmt.Printf("✓ Found %d tables\n", len(tables))

	fmt.Println("\n✓ All tests passed! Package import is working correctly.")
	fmt.Println("\nPackage structure:")
	fmt.Println("- github.com/tanqiangyes/go-word/pkg/word - Main word processing package")
	fmt.Println("- github.com/tanqiangyes/go-word/pkg/types - Shared type definitions")
	fmt.Println("- github.com/tanqiangyes/go-word/pkg/opc - OPC container handling")
	fmt.Println("- github.com/tanqiangyes/go-word/pkg/parser - XML parsing")
}
