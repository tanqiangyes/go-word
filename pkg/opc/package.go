// Package opc provides Open Packaging Convention (OPC) container functionality
// for handling Word documents and other Office Open XML formats.
package opc

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"path"
	"strings"
)

// Container represents an OPC container (ZIP-based package)
type Container struct {
	reader *zip.ReadCloser
	writer *zip.Writer
	buffer *bytes.Buffer
	parts  map[string]*Part
}

// Part represents a part within the OPC container
type Part struct {
	Name     string
	Content  []byte
	ContentType string
}

// Relationship represents a relationship between parts
type Relationship struct {
	ID     string
	Type   string
	Target string
}

// Open opens an OPC container from a file
func Open(filename string) (*Container, error) {
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open OPC container: %w", err)
	}
	
	return &Container{
		reader: reader,
		parts:  make(map[string]*Part),
	}, nil
}

// OpenFromReader opens an OPC container from an io.Reader
func OpenFromReader(r io.Reader) (*Container, error) {
	data, err := io.ReadAll(r)
	if err != nil {
		return nil, fmt.Errorf("failed to read container data: %w", err)
	}
	
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("failed to create zip reader: %w", err)
	}
	
	return &Container{
		reader: &zip.ReadCloser{
			Reader: *reader,
		},
		parts: make(map[string]*Part),
	}, nil
}

// Close closes the container and releases resources
func (c *Container) Close() error {
	if c.reader != nil {
		return c.reader.Close()
	}
	return nil
}

// GetPart retrieves a part by name
func (c *Container) GetPart(name string) (*Part, error) {
	if c.reader == nil {
		return nil, fmt.Errorf("container not opened for reading")
	}
	
	for _, file := range c.reader.File {
		if file.Name == name {
			rc, err := file.Open()
			if err != nil {
				return nil, fmt.Errorf("failed to open part %s: %w", name, err)
			}
			defer rc.Close()
			
			content, err := io.ReadAll(rc)
			if err != nil {
				return nil, fmt.Errorf("failed to read part %s: %w", name, err)
			}
			
			return &Part{
				Name:     name,
				Content:  content,
				ContentType: getContentType(name),
			}, nil
		}
	}
	
	return nil, fmt.Errorf("part not found: %s", name)
}

// ListParts returns all parts in the container
func (c *Container) ListParts() ([]string, error) {
	if c.reader == nil {
		return nil, fmt.Errorf("container not opened for reading")
	}
	
	var parts []string
	for _, file := range c.reader.File {
		parts = append(parts, file.Name)
	}
	
	return parts, nil
}

// GetRelationships retrieves relationships for a given part
func (c *Container) GetRelationships(partName string) ([]Relationship, error) {
	relsPath := getRelationshipsPath(partName)
	relsPart, err := c.GetPart(relsPath)
	if err != nil {
		return nil, fmt.Errorf("failed to get relationships for %s: %w", partName, err)
	}
	
	return parseRelationships(relsPart.Content)
}

// getContentType determines the content type based on file extension
func getContentType(filename string) string {
	ext := strings.ToLower(path.Ext(filename))
	switch ext {
	case ".xml":
		return "application/xml"
	case ".rels":
		return "application/vnd.openxmlformats-package.relationships+xml"
	case ".png":
		return "image/png"
	case ".jpg", ".jpeg":
		return "image/jpeg"
	case ".gif":
		return "image/gif"
	default:
		return "application/octet-stream"
	}
}

// getRelationshipsPath returns the path to the relationships file for a part
func getRelationshipsPath(partName string) string {
	dir := path.Dir(partName)
	if dir == "." {
		return "_rels/.rels"
	}
	return path.Join(dir, "_rels", path.Base(partName)+".rels")
}

// parseRelationships parses the relationships XML content
func parseRelationships(content []byte) ([]Relationship, error) {
	// TODO: Implement XML parsing for relationships
	// For now, return empty slice
	return []Relationship{}, nil
}

// AddPart adds a part to the container
func (c *Container) AddPart(name string, content []byte, contentType string) {
	if c.parts == nil {
		c.parts = make(map[string]*Part)
	}
	
	c.parts[name] = &Part{
		Name:        name,
		Content:     content,
		ContentType: contentType,
	}
}

// SaveToFile saves the container to a file
func (c *Container) SaveToFile(filename string) error {
	if c.parts == nil || len(c.parts) == 0 {
		return fmt.Errorf("no parts to save")
	}

	file, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	writer := zip.NewWriter(file)
	defer writer.Close()

	for name, part := range c.parts {
		zipFile, err := writer.Create(name)
		if err != nil {
			return fmt.Errorf("failed to create zip entry %s: %w", name, err)
		}

		_, err = zipFile.Write(part.Content)
		if err != nil {
			return fmt.Errorf("failed to write part %s: %w", name, err)
		}
	}

	return nil
} 