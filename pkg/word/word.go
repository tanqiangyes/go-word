// Package word provides a comprehensive Word document processing library.
// This package allows you to create, read, modify, and save Word documents (.docx files).
//
// The package is designed to be easy to use while providing powerful features:
//   - Document creation and manipulation
//   - Text and formatting operations
//   - Table and image support
//   - Style management
//   - Document validation and protection
//
// Example usage:
//
//	package main
//
//	import (
//		"fmt"
//		"log"
//		"github.com/tanqiangyes/go-word/pkg/word"
//	)
//
//	func main() {
//		// Open an existing document
//		doc, err := word.Open("document.docx")
//		if err != nil {
//			log.Fatal(err)
//		}
//		defer doc.Close()
//
//		// Get document text
//		text, err := doc.GetText()
//		if err != nil {
//			log.Fatal(err)
//		}
//		fmt.Println("Document content:", text)
//
//		// Create a new document
//		newDoc, err := word.New()
//		if err != nil {
//			log.Fatal(err)
//		}
//		defer newDoc.Close()
//	}
package word

// This package exports the following main functions:
//   - word.Open(filename string) (*Document, error) - Open an existing document
//   - word.New() (*Document, error) - Create a new document
//
// And the following main types:
//   - Document - Main document interface
//   - Paragraph, Run, Table, TableRow, TableCell - Content structures
//   - Style, StyleProperties, Font, Color - Formatting properties
//
// See individual files for detailed API documentation.
