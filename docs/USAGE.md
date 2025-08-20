# Go-Word Package Usage Guide

This guide explains how to use the Go-Word package in your own projects.

## Installation

### Using Go Modules (Recommended)

If your project uses Go modules (go.mod), simply add the dependency:

```bash
go get github.com/tanqiangyes/go-word
```

This will automatically add the dependency to your `go.mod` file.

### Manual Installation

```bash
git clone https://github.com/tanqiangyes/go-word.git
cd go-word
go mod tidy
```

## Basic Usage

### Importing the Package

```go
import (
    "github.com/tanqiangyes/go-word/pkg/word"
    "github.com/tanqiangyes/go-word/pkg/types"
)
```

### Opening Existing Documents

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    // Open an existing Word document
    doc, err := word.Open("path/to/document.docx")
    if err != nil {
        log.Fatal("Failed to open document:", err)
    }
    defer doc.Close() // Always close the document when done

    // Extract text content
    text, err := doc.GetText()
    if err != nil {
        log.Fatal("Failed to get text:", err)
    }
    fmt.Println("Document content:", text)

    // Get all paragraphs
    paragraphs, err := doc.GetParagraphs()
    if err != nil {
        log.Fatal("Failed to get paragraphs:", err)
    }
    
    for i, paragraph := range paragraphs {
        fmt.Printf("Paragraph %d: %s\n", i+1, paragraph.Text)
        
        // Access individual runs (text with formatting)
        for j, run := range paragraph.Runs {
            fmt.Printf("  Run %d: '%s' (Bold: %v, Italic: %v)\n",
                j+1, run.Text, run.Bold, run.Italic)
        }
    }

    // Get all tables
    tables, err := doc.GetTables()
    if err != nil {
        log.Fatal("Failed to get tables:", err)
    }
    
    for i, table := range tables {
        fmt.Printf("Table %d: %d rows x %d columns\n",
            i+1, len(table.Rows), table.Columns)
        
        // Access table content
        for rowIdx, row := range table.Rows {
            for colIdx, cell := range row.Cells {
                fmt.Printf("  Cell [%d,%d]: %s\n", rowIdx, colIdx, cell.Text)
            }
        }
    }
}
```

### Creating New Documents

```go
package main

import (
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
    "github.com/tanqiangyes/go-word/pkg/types"
)

func main() {
    // Create a new empty document
    doc, err := word.New()
    if err != nil {
        log.Fatal("Failed to create document:", err)
    }
    defer doc.Close()

    // The document is now ready for content addition
    // Note: Content addition methods are coming in future versions
    
    fmt.Println("New document created successfully!")
}
```

### Working with Document Types

The package provides several types for working with document content:

```go
import "github.com/tanqiangyes/go-word/pkg/types"

// Create a paragraph
paragraph := types.Paragraph{
    Text: "This is a sample paragraph",
    Runs: []types.Run{
        {
            Text: "This is ",
            Bold: true,
            FontSize: 14,
        },
        {
            Text: "formatted text",
            Italic: true,
            FontSize: 16,
        },
    },
}

// Create a table
table := types.Table{
    Columns: 3,
    Rows: []types.TableRow{
        {
            Cells: []types.TableCell{
                {Text: "Header 1"},
                {Text: "Header 2"},
                {Text: "Header 3"},
            },
        },
        {
            Cells: []types.TableCell{
                {Text: "Data 1"},
                {Text: "Data 2"},
                {Text: "Data 3"},
            },
        },
    },
}

// Create a style
style := types.Style{
    Name: "CustomStyle",
    Type: types.ParagraphStyle,
    Properties: &types.StyleProperties{
        FontName: "Arial",
        FontSize: 12,
        Bold:     true,
        Color:    &types.Color{Value: "#000000"},
    },
}
```

## Advanced Features

### Document Parts

Access different parts of the document:

```go
// Get document parts manager
parts := doc.GetDocumentParts()

// Access specific parts (headers, footers, styles, etc.)
// Implementation details coming in future versions
```

### Document Validation

```go
// Validate document structure
// Implementation details coming in future versions
```

### Document Protection

```go
// Set document protection
// Implementation details coming in future versions
```

## Error Handling

The package follows Go's error handling patterns:

```go
doc, err := word.Open("document.docx")
if err != nil {
    // Handle specific error types
    switch {
    case os.IsNotExist(err):
        log.Fatal("Document file not found")
    case strings.Contains(err.Error(), "corrupted"):
        log.Fatal("Document appears to be corrupted")
    default:
        log.Fatal("Failed to open document:", err)
    }
}
```

## Best Practices

1. **Always close documents**: Use `defer doc.Close()` after opening a document
2. **Check errors**: Always check returned errors from package functions
3. **Use appropriate types**: Use the types from `pkg/types` for creating content
4. **Handle resources**: Be mindful of memory usage when working with large documents

## Troubleshooting

### Common Issues

1. **Import errors**: Make sure you're using the correct import path
2. **File not found**: Verify the document path is correct
3. **Permission denied**: Ensure you have read/write permissions for the file
4. **Corrupted documents**: Some documents may have unsupported features

### Getting Help

- Check the examples in the `examples/` directory
- Review the package documentation
- Open an issue on GitHub for bugs or feature requests

## Next Steps

This package is actively developed. Future versions will include:

- Content addition and modification
- Document saving
- Advanced formatting options
- Template support
- Performance optimizations

Stay tuned for updates!
