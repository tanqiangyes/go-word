# Go-Word

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
go get github.com/tanqiangyes/go-word
```

## Quick Start

### Basic Usage

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
    "github.com/tanqiangyes/go-word/pkg/types"
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

// Create content using types from the types package
paragraph := types.Paragraph{
    Text: "Hello, World!",
    Runs: []types.Run{
        {
            Text: "Hello, ",
            Bold: true,
        },
        {
            Text: "World!",
            Italic: true,
        },
    },
}

// The document is now ready for content addition
// (Implementation details for adding content coming soon)
```

## Package Structure

This project is organized into several packages:

- **`pkg/word`** - Main word processing package with document operations
- **`pkg/types`** - Shared type definitions for documents, paragraphs, tables, etc.
- **`pkg/opc`** - Open Packaging Convention container handling
- **`pkg/parser`** - XML parsing for WordprocessingML
- **`pkg/utils`** - Utility functions and logging
- **`pkg/plugin`** - Plugin system for extending functionality

## API Reference

### Main Functions (pkg/word)

- `word.Open(filename string) (*Document, error)` - Open an existing document
- `word.New() (*Document, error)` - Create a new document

### Document Methods

- `doc.GetText() (string, error)` - Get plain text content
- `doc.GetParagraphs() ([]Paragraph, error)` - Get all paragraphs
- `doc.GetTables() ([]Table, error)` - Get all tables
- `doc.GetDocumentParts() *DocumentParts` - Access document parts
- `doc.Close() error` - Close and release resources

### Types (pkg/types)

- `Document` - Main document interface
- `Paragraph` - Document paragraph
- `Run` - Text run with formatting
- `Table`, `TableRow`, `TableCell` - Table structures
- `Style`, `StyleProperties` - Document styles
- `Font`, `Color` - Formatting properties

## Examples

See the `examples/` directory for detailed usage examples:

- `examples/basic_usage/` - Basic document operations
- `examples/import_test/` - Package import and type usage testing
- `examples/advanced_styles/` - Style management examples
- `examples/templates/` - Document templating

## Development

### Building

#### 使用构建脚本 (推荐)

我们提供了智能构建脚本来处理不同环境的构建需求：

**Linux/macOS/WSL:**
```bash
# 只构建核心包 (默认，推荐用于库开发)
./build.sh --core-only

# 构建核心包和示例
./build.sh --with-examples

# 构建所有包 (包括GUI，需要X11支持)
./build.sh --all

# 查看所有选项
./build.sh --help
```

**Windows:**
```cmd
# 只构建核心包 (默认，推荐用于库开发)
build.bat --core-only

# 构建核心包和示例
build.bat --with-examples

# 构建所有包
build.bat --all

# 查看所有选项
build.bat --help
```

#### 手动构建

```bash
# Build all packages
go build ./...

# Build specific package
go build ./pkg/word

# Build examples
go build ./examples/basic_usage/
```

**注意**: 在WSL环境中，GUI包可能需要额外的X11库支持。如果遇到编译错误，请使用 `--core-only` 选项。

### Testing

```bash
# Run all tests
go test ./...

# Run tests with coverage
go test -cover ./...

# Run specific package tests
go test ./pkg/word/...
```

### Running Examples

```bash
# Build and run basic usage example
go build -o examples/basic_usage/basic_usage ./examples/basic_usage/
./examples/basic_usage/basic_usage

# Build and run import test
go build -o examples/import_test/import_test ./examples/import_test/
./examples/import_test/import_test
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 