package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewDocumentParts(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	if parts == nil {
		t.Fatal("Expected DocumentParts to be created")
	}
	
	if parts.MainDocumentPart == nil {
		t.Error("Expected MainDocumentPart to be initialized")
	}
}

func TestAddHeaderPart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	headerPart := wordprocessingml.HeaderPart{
		ID:      "header1",
		Type:    wordprocessingml.HeaderType,
		Content: []types.Paragraph{},
	}
	
	parts.AddHeaderPart(headerPart)
	
	if len(parts.HeaderParts) != 1 {
		t.Errorf("Expected 1 header part, got %d", len(parts.HeaderParts))
	}
	
	if parts.HeaderParts[0].ID != "header1" {
		t.Errorf("Expected header ID 'header1', got '%s'", parts.HeaderParts[0].ID)
	}
}

func TestAddFooterPart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	footerPart := wordprocessingml.FooterPart{
		ID:      "footer1",
		Type:    wordprocessingml.FooterType,
		Content: []types.Paragraph{},
	}
	
	parts.AddFooterPart(footerPart)
	
	if len(parts.FooterParts) != 1 {
		t.Errorf("Expected 1 footer part, got %d", len(parts.FooterParts))
	}
	
	if parts.FooterParts[0].ID != "footer1" {
		t.Errorf("Expected footer ID 'footer1', got '%s'", parts.FooterParts[0].ID)
	}
}

func TestAddCommentPart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	commentPart := wordprocessingml.CommentPart{
		ID:      "comment1",
		Content: []wordprocessingml.Comment{},
	}
	
	parts.AddCommentPart(commentPart)
	
	if len(parts.CommentParts) != 1 {
		t.Errorf("Expected 1 comment part, got %d", len(parts.CommentParts))
	}
	
	if parts.CommentParts[0].ID != "comment1" {
		t.Errorf("Expected comment ID 'comment1', got '%s'", parts.CommentParts[0].ID)
	}
}

func TestAddFootnotePart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	footnotePart := wordprocessingml.FootnotePart{
		ID:      "footnote1",
		Content: []wordprocessingml.Footnote{},
	}
	
	parts.AddFootnotePart(footnotePart)
	
	if len(parts.FootnoteParts) != 1 {
		t.Errorf("Expected 1 footnote part, got %d", len(parts.FootnoteParts))
	}
	
	if parts.FootnoteParts[0].ID != "footnote1" {
		t.Errorf("Expected footnote ID 'footnote1', got '%s'", parts.FootnoteParts[0].ID)
	}
}

func TestAddEndnotePart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	endnotePart := wordprocessingml.EndnotePart{
		ID:      "endnote1",
		Content: []wordprocessingml.Endnote{},
	}
	
	parts.AddEndnotePart(endnotePart)
	
	if len(parts.EndnoteParts) != 1 {
		t.Errorf("Expected 1 endnote part, got %d", len(parts.EndnoteParts))
	}
	
	if parts.EndnoteParts[0].ID != "endnote1" {
		t.Errorf("Expected endnote ID 'endnote1', got '%s'", parts.EndnoteParts[0].ID)
	}
}

func TestGetHeaderPart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	headerPart := wordprocessingml.HeaderPart{
		ID:      "header1",
		Type:    wordprocessingml.HeaderType,
		Content: []types.Paragraph{},
	}
	
	parts.AddHeaderPart(headerPart)
	
	found := parts.GetHeaderPart("header1")
	if found == nil {
		t.Error("Expected to find header part")
	}
	
	if found.ID != "header1" {
		t.Errorf("Expected header ID 'header1', got '%s'", found.ID)
	}
}

func TestGetFooterPart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	footerPart := wordprocessingml.FooterPart{
		ID:      "footer1",
		Type:    wordprocessingml.FooterType,
		Content: []types.Paragraph{},
	}
	
	parts.AddFooterPart(footerPart)
	
	found := parts.GetFooterPart("footer1")
	if found == nil {
		t.Error("Expected to find footer part")
	}
	
	if found.ID != "footer1" {
		t.Errorf("Expected footer ID 'footer1', got '%s'", found.ID)
	}
}

func TestGetCommentPart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	commentPart := wordprocessingml.CommentPart{
		ID:      "comment1",
		Content: []wordprocessingml.Comment{},
	}
	
	parts.AddCommentPart(commentPart)
	
	found := parts.GetCommentPart("comment1")
	if found == nil {
		t.Error("Expected to find comment part")
	}
	
	if found.ID != "comment1" {
		t.Errorf("Expected comment ID 'comment1', got '%s'", found.ID)
	}
}

func TestGetFootnotePart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	footnotePart := wordprocessingml.FootnotePart{
		ID:      "footnote1",
		Content: []wordprocessingml.Footnote{},
	}
	
	parts.AddFootnotePart(footnotePart)
	
	found := parts.GetFootnotePart("footnote1")
	if found == nil {
		t.Error("Expected to find footnote part")
	}
	
	if found.ID != "footnote1" {
		t.Errorf("Expected footnote ID 'footnote1', got '%s'", found.ID)
	}
}

func TestGetEndnotePart(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	endnotePart := wordprocessingml.EndnotePart{
		ID:      "endnote1",
		Content: []wordprocessingml.Endnote{},
	}
	
	parts.AddEndnotePart(endnotePart)
	
	found := parts.GetEndnotePart("endnote1")
	if found == nil {
		t.Error("Expected to find endnote part")
	}
	
	if found.ID != "endnote1" {
		t.Errorf("Expected endnote ID 'endnote1', got '%s'", found.ID)
	}
}

func TestGetPartsSummary(t *testing.T) {
	parts := wordprocessingml.NewDocumentParts()
	
	// 添加各种部分
	parts.AddHeaderPart(wordprocessingml.HeaderPart{ID: "header1", Type: wordprocessingml.HeaderType, Content: []types.Paragraph{}})
	parts.AddFooterPart(wordprocessingml.FooterPart{ID: "footer1", Type: wordprocessingml.FooterType, Content: []types.Paragraph{}})
	parts.AddCommentPart(wordprocessingml.CommentPart{ID: "comment1", Content: []wordprocessingml.Comment{}})
	parts.AddFootnotePart(wordprocessingml.FootnotePart{ID: "footnote1", Content: []wordprocessingml.Footnote{}})
	parts.AddEndnotePart(wordprocessingml.EndnotePart{ID: "endnote1", Content: []wordprocessingml.Endnote{}})
	
	summary := parts.GetPartsSummary()
	
	if summary == "" {
		t.Error("Expected non-empty parts summary")
	}
	
	// 检查摘要是否包含预期的部分信息
	expectedParts := []string{"MainDocumentPart", "HeaderParts: 1", "FooterParts: 1", "CommentParts: 1", "FootnoteParts: 1", "EndnoteParts: 1"}
	for _, expected := range expectedParts {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

// 辅助函数
