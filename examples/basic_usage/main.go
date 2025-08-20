package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
	fmt.Println("Go-Word Package Test")
	fmt.Println("====================")

	// Test creating a new document
	fmt.Println("\n1. Creating a new document...")
	doc, err := word.New()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}
	defer doc.Close()

	fmt.Println("✓ Document created successfully")

	// Test getting document parts summary
	fmt.Println("\n2. Getting document parts summary...")
	summary := doc.GetPartsSummary()
	fmt.Println("Document summary:", summary)

	// Test getting paragraphs (should be empty for new document)
	fmt.Println("\n3. Getting paragraphs...")
	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		log.Fatal("Failed to get paragraphs:", err)
	}
	fmt.Printf("✓ Found %d paragraphs\n", len(paragraphs))

	// Test getting tables (should be empty for new document)
	fmt.Println("\n4. Getting tables...")
	tables, err := doc.GetTables()
	if err != nil {
		log.Fatal("Failed to get tables:", err)
	}
	fmt.Printf("✓ Found %d tables\n", len(tables))

	fmt.Println("\n✓ All tests passed! Package is working correctly.")
}
