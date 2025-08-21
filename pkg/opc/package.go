// Package opc provides Open Packaging Convention (OPC) container functionality
// for handling Word documents and other Office Open XML formats.
package opc

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"strings"
)

// Container represents an OPC container (ZIP-based package)
type Container struct {
	Reader *zip.Reader
	Writer *zip.Writer
	Buffer *bytes.Buffer
	Parts  map[string]*Part
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

// New creates a new empty OPC container
func New() (*Container, error) {
	return &Container{
		Parts: make(map[string]*Part),
	}, nil
}

// Open opens an OPC container from a file
func Open(filename string) (*Container, error) {
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open OPC container: %w", err)
	}
	defer reader.Close()
	
	// Convert to zip.Reader for the container
	zipReader := &reader.Reader
	
	return &Container{
		Reader: zipReader,
		Parts:  make(map[string]*Part),
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
		Reader: reader,
		Parts: make(map[string]*Part),
	}, nil
}

// Close closes the container and releases resources
func (c *Container) Close() error {
	// zip.Reader doesn't have a Close method, so we just clean up references
	c.Reader = nil
	c.Writer = nil
	c.Buffer = nil
	c.Parts = nil
	return nil
}

// GetPart retrieves a part by name
func (c *Container) GetPart(name string) (*Part, error) {
	if c.Reader == nil {
		return nil, fmt.Errorf("container not opened for reading")
	}
	
	for _, file := range c.Reader.File {
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
	if c.Reader == nil {
		return nil, fmt.Errorf("container not opened for reading")
	}
	
	var parts []string
	for _, file := range c.Reader.File {
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

// RelationshipsXML represents the XML structure for relationships
type RelationshipsXML struct {
	XMLName       xml.Name           `xml:"Relationships"`
	Namespace     string             `xml:"xmlns,attr"`
	Relationships []RelationshipXML  `xml:"Relationship"`
}

// RelationshipXML represents a single relationship in XML
type RelationshipXML struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// parseRelationships parses the relationships XML content
func parseRelationships(content []byte) ([]Relationship, error) {
	var relsXML RelationshipsXML
	err := xml.Unmarshal(content, &relsXML)
	if err != nil {
		return nil, fmt.Errorf("failed to parse relationships XML: %w", err)
	}
	
	var relationships []Relationship
	for _, rel := range relsXML.Relationships {
		relationships = append(relationships, Relationship{
			ID:     rel.ID,
			Type:   rel.Type,
			Target: rel.Target,
		})
	}
	
	return relationships, nil
}

// AddPart adds a part to the container
func (c *Container) AddPart(name string, content []byte, contentType string) {
	if c.Parts == nil {
		c.Parts = make(map[string]*Part)
	}
	
	c.Parts[name] = &Part{
		Name:        name,
		Content:     content,
		ContentType: contentType,
	}
}

// SaveToFile saves the container to a file
func (c *Container) SaveToFile(filename string) error {
	if c.Parts == nil || len(c.Parts) == 0 {
		return fmt.Errorf("no parts to save")
	}

	file, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	writer := zip.NewWriter(file)
	defer writer.Close()

	for name, part := range c.Parts {
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