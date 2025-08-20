# Go-Word Package

A comprehensive Go library for processing Word documents (.docx files) according to the Office Open XML specification.

## Features

- **Document Creation**: Create new Word documents from scratch
- **Document Reading**: Open and parse existing .docx files
- **Content Extraction**: Extract text, paragraphs, tables, and formatting
- **Document Manipulation**: Modify content, styles, and structure
- **Advanced Features**: Support for images, charts, and complex formatting
- **Style Management**: Comprehensive style system with inheritance
- **Document Protection**: Password protection and permission controls
- **Validation**: Document structure and content validation

## Installation

```bash
go get github.com/tanqiangyes/go-word/pkg/word
```

## Quick Start

### Basic Usage

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // Open an existing document
    doc, err := word.Open("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // Get document text
    text, err := doc.GetText()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Println("Document content:", text)

    // Get paragraphs
    paragraphs, err := doc.GetParagraphs()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Printf("Found %d paragraphs\n", len(paragraphs))

    // Get tables
    tables, err := doc.GetTables()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Printf("Found %d tables\n", len(tables))
}
```

### Creating New Documents

```go
// Create a new empty document
doc, err := word.New()
if err != nil {
    log.Fatal(err)
}
defer doc.Close()

// The document is now ready for content addition
// (Implementation details for adding content coming soon)
```

## API Reference

### Main Functions

- `word.Open(filename string) (*Document, error)` - Open an existing document
- `word.New() (*Document, error)` - Create a new document

### Document Methods

- `doc.GetText() (string, error)` - Get plain text content
- `doc.GetParagraphs() ([]Paragraph, error)` - Get all paragraphs
- `doc.GetTables() ([]Table, error)` - Get all tables
- `doc.GetDocumentParts() *DocumentParts` - Access document parts
- `doc.Close() error` - Close and release resources

### Types

The package exports the following types for convenience:

- `Document` - Main document interface
- `Paragraph` - Document paragraph
- `Run` - Text run with formatting
- `Table`, `TableRow`, `TableCell` - Table structures
- `Style`, `StyleProperties` - Document styles
- `Font`, `Color` - Formatting properties

## Examples

See the `examples/` directory for more detailed usage examples:

- `examples/basic_usage/` - Basic document operations
- `examples/advanced_styles/` - Style management examples
- `examples/templates/` - Document templating

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
