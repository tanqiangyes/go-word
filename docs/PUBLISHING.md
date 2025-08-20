# Publishing Go-Word as a Go Module

This guide explains how to publish and use the Go-Word package as a Go module.

## Package Structure

The Go-Word project is organized into several packages that can be imported individually:

### Core Packages

- **`github.com/tanqiangyes/go-word/pkg/word`** - Main word processing package
  - `word.Open(filename string) (*Document, error)` - Open existing documents
  - `word.New() (*Document, error)` - Create new documents
  - `Document` struct with methods for content access

- **`github.com/tanqiangyes/go-word/pkg/types`** - Shared type definitions
  - `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`
  - `Style`, `StyleProperties`, `Font`, `Color`
  - `DocumentContent`, `CoreProperties`

- **`github.com/tanqiangyes/go-word/pkg/opc`** - OPC container handling
  - `Container` struct for managing document parts
  - `Open(filename string) (*Container, error)`
  - `New() (*Container, error)`

- **`github.com/tanqiangyes/go-word/pkg/parser`** - XML parsing
  - `ParseWordML(data []byte) (*types.DocumentContent, error)`

## Installation

### Using Go Modules

```bash
# Add the dependency to your project
go get github.com/tanqiangyes/go-word

# This will add to your go.mod:
# require github.com/tanqiangyes/go-word v0.0.0-20240101000000-000000000000
```

### Import in Your Code

```go
package main

import (
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

    // Get document content
    text, err := doc.GetText()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Println("Document content:", text)
}
```

## Module Information

### go.mod

```go
module github.com/tanqiangyes/go-word

go 1.22

require fyne.io/fyne/v2 v2.4.3

// ... other dependencies
```

### Module Path

The module path `github.com/tanqiangyes/go-word` follows Go's standard convention:
- `github.com/tanqiangyes` - GitHub username/organization
- `go-word` - Repository name

### Versioning

Currently, the package is at version `v0.0.0` (pre-release). For production use, consider:

1. **Tagging releases**: Create git tags for stable versions
2. **Semantic versioning**: Follow semver.org guidelines
3. **Release notes**: Document changes between versions

## Publishing Process

### 1. Prepare the Repository

```bash
# Ensure all tests pass
go test ./...

# Build all packages
go build ./pkg/...

# Check for any compilation errors
go mod tidy
```

### 2. Create a Release Tag

```bash
# Create and push a tag
git tag v0.1.0
git push origin v0.1.0
```

### 3. Verify Module Availability

```bash
# Test module download
go get github.com/tanqiangyes/go-word@v0.1.0

# Or test in a new project
mkdir test-project
cd test-project
go mod init test-project
go get github.com/tanqiangyes/go-word
```

## Usage Examples

### Basic Document Reading

```go
package main

import (
    "fmt"
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    doc, err := word.Open("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    text, err := doc.GetText()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Println(text)
}
```

### Working with Types

```go
package main

import (
    "github.com/tanqiangyes/go-word/pkg/types"
)

func createContent() {
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

    table := types.Table{
        Columns: 2,
        Rows: []types.TableRow{
            {
                Cells: []types.TableCell{
                    {Text: "Name"},
                    {Text: "Age"},
                },
            },
            {
                Cells: []types.TableCell{
                    {Text: "John"},
                    {Text: "30"},
                },
            },
        },
    }
}
```

### Creating New Documents

```go
package main

import (
    "log"
    "github.com/tanqiangyes/go-word/pkg/word"
)

func main() {
    doc, err := word.New()
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    // Document is ready for content addition
    // (Implementation coming in future versions)
}
```

## Troubleshooting

### Common Issues

1. **Import errors**: Ensure you're using the correct import path
2. **Version conflicts**: Use `go mod tidy` to resolve dependencies
3. **Build errors**: Check that all required packages are available

### Getting Help

- Check the examples in the `examples/` directory
- Review the package documentation
- Open an issue on GitHub for bugs or feature requests

## Future Enhancements

The package is actively developed. Planned features include:

- Content addition and modification methods
- Document saving functionality
- Advanced formatting options
- Template support
- Performance optimizations
- Extended format support

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.
