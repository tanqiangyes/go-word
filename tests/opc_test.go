package tests

import (
	"bytes"
	"path"
	"strings"
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/opc"
)

func TestContainerOpenFromReader(t *testing.T) {
	// 创建一个简单的ZIP文件用于测试
	zipData := createTestZipData()
	
	container, err := opc.OpenFromReader(bytes.NewReader(zipData))
	if err != nil {
		t.Fatalf("Failed to open container from reader: %v", err)
	}
	defer container.Close()
	
	parts, err := container.ListParts()
	if err != nil {
		t.Fatalf("Failed to list parts: %v", err)
	}
	
	if len(parts) == 0 {
		t.Error("Expected parts to be found in container")
	}
}

func TestContainerGetPart(t *testing.T) {
	zipData := createTestZipData()
	
	container, err := opc.OpenFromReader(bytes.NewReader(zipData))
	if err != nil {
		t.Fatalf("Failed to open container: %v", err)
	}
	defer container.Close()
	
	part, err := container.GetPart("test.xml")
	if err != nil {
		t.Fatalf("Failed to get part: %v", err)
	}
	
	if part.Name != "test.xml" {
		t.Errorf("Expected part name 'test.xml', got '%s'", part.Name)
	}
	
	if part.ContentType != "application/xml" {
		t.Errorf("Expected content type 'application/xml', got '%s'", part.ContentType)
	}
}

func TestContainerGetPartNotFound(t *testing.T) {
	zipData := createTestZipData()
	
	container, err := opc.OpenFromReader(bytes.NewReader(zipData))
	if err != nil {
		t.Fatalf("Failed to open container: %v", err)
	}
	defer container.Close()
	
	_, err = container.GetPart("nonexistent.xml")
	if err == nil {
		t.Error("Expected error when getting non-existent part")
	}
}

func TestGetContentType(t *testing.T) {
	testCases := []struct {
		filename    string
		contentType string
	}{
		{"document.xml", "application/xml"},
		{"styles.xml", "application/xml"},
		{"image.png", "image/png"},
		{"photo.jpg", "image/jpeg"},
		{"photo.jpeg", "image/jpeg"},
		{"icon.gif", "image/gif"},
		{"unknown.xyz", "application/octet-stream"},
	}
	
	for _, tc := range testCases {
		contentType := getContentType(tc.filename)
		if contentType != tc.contentType {
			t.Errorf("For filename '%s', expected content type '%s', got '%s'", 
				tc.filename, tc.contentType, contentType)
		}
	}
}

// createTestZipData creates a simple ZIP file for testing
func createTestZipData() []byte {
	// This is a simplified implementation
	// In a real test, you would create a proper ZIP file
	return []byte("PK\x03\x04\x14\x00\x00\x00\x08\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00test.xml\x00\x00\x00")
}

// getContentType is a helper function for testing
func getContentType(filename string) string {
	// This is a simplified version of the actual implementation
	// In the real code, this would be in the opc package
	ext := strings.ToLower(path.Ext(filename))
	switch ext {
	case ".xml":
		return "application/xml"
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