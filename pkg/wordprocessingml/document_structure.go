// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"sort"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// DocumentStructure represents the structure of a Word document
type DocumentStructure struct {
	Document *Document
	Sections []DocumentSection
	Outline  DocumentOutline
}

// DocumentSection represents a section of the document
type DocumentSection struct {
	ID          string
	Title       string
	Level       int
	StartIndex  int
	EndIndex    int
	Paragraphs  []types.Paragraph
	SubSections []DocumentSection
}

// DocumentOutline represents the document outline
type DocumentOutline struct {
	Title    string
	Sections []DocumentSection
	Level    int
}

// ReorganizeDocument reorganizes the document structure
func (d *Document) ReorganizeDocument() (*DocumentStructure, error) {
	if d.mainPart == nil || d.mainPart.Content == nil {
		return nil, fmt.Errorf("document content not loaded")
	}

	structure := &DocumentStructure{
		Document: d,
		Sections: make([]DocumentSection, 0),
	}

	// 分析文档结构
	if err := structure.analyzeDocumentStructure(); err != nil {
		return nil, fmt.Errorf("failed to analyze document structure: %w", err)
	}

	// 生成文档大纲
	if err := structure.generateOutline(); err != nil {
		return nil, fmt.Errorf("failed to generate outline: %w", err)
	}

	return structure, nil
}

// analyzeDocumentStructure analyzes the document structure
func (ds *DocumentStructure) analyzeDocumentStructure() error {
	paragraphs := ds.Document.mainPart.Content.Paragraphs
	currentSection := DocumentSection{
		ID:         "section_1",
		Title:      "默认段落",
		Level:      1,
		StartIndex: 0,
		Paragraphs: make([]types.Paragraph, 0),
	}

	for i, paragraph := range paragraphs {
		// 检查是否是标题段落
		if isHeadingParagraph(paragraph) {
			// 结束当前段落
			if len(currentSection.Paragraphs) > 0 {
				currentSection.EndIndex = i - 1
				ds.Sections = append(ds.Sections, currentSection)
			}

			// 开始新段落
			currentSection = DocumentSection{
				ID:         fmt.Sprintf("section_%d", len(ds.Sections)+1),
				Title:      extractHeadingTitle(paragraph),
				Level:      determineHeadingLevel(paragraph),
				StartIndex: i,
				Paragraphs: make([]types.Paragraph, 0),
			}
		}

		currentSection.Paragraphs = append(currentSection.Paragraphs, paragraph)
	}

	// 添加最后一个段落
	if len(currentSection.Paragraphs) > 0 {
		currentSection.EndIndex = len(paragraphs) - 1
		ds.Sections = append(ds.Sections, currentSection)
	}

	return nil
}

// generateOutline generates the document outline
func (ds *DocumentStructure) generateOutline() error {
	ds.Outline = DocumentOutline{
		Title:    "文档大纲",
		Sections: make([]DocumentSection, 0),
		Level:    0,
	}

	// 构建层次结构
	sectionMap := make(map[int]*DocumentSection)
	
	for i := range ds.Sections {
		section := &ds.Sections[i]
		sectionMap[i] = section
		
		// 查找父段落
		for j := i - 1; j >= 0; j-- {
			parent := sectionMap[j]
			if parent.Level < section.Level {
				parent.SubSections = append(parent.SubSections, *section)
				break
			}
		}
	}

	// 添加顶级段落到大纲
	for _, section := range ds.Sections {
		if section.Level == 1 {
			ds.Outline.Sections = append(ds.Outline.Sections, section)
		}
	}

	return nil
}

// ReorderContent reorders the document content
func (ds *DocumentStructure) ReorderContent(newOrder []string) error {
	if len(newOrder) != len(ds.Sections) {
		return fmt.Errorf("new order length (%d) doesn't match sections count (%d)", 
			len(newOrder), len(ds.Sections))
	}

	// 创建新的段落顺序
	newParagraphs := make([]types.Paragraph, 0)
	
	// 根据新顺序重新排列段落
	for _, sectionID := range newOrder {
		for _, section := range ds.Sections {
			if section.ID == sectionID {
				newParagraphs = append(newParagraphs, section.Paragraphs...)
				break
			}
		}
	}

	// 更新文档内容
	ds.Document.mainPart.Content.Paragraphs = newParagraphs
	
	// 重新分析结构
	return ds.analyzeDocumentStructure()
}

// GetOutline returns the document outline
func (ds *DocumentStructure) GetOutline() DocumentOutline {
	return ds.Outline
}

// GetSections returns all document sections
func (ds *DocumentStructure) GetSections() []DocumentSection {
	return ds.Sections
}

// FindSectionByTitle finds a section by its title
func (ds *DocumentStructure) FindSectionByTitle(title string) (*DocumentSection, error) {
	for i := range ds.Sections {
		if ds.Sections[i].Title == title {
			return &ds.Sections[i], nil
		}
	}
	return nil, fmt.Errorf("section with title '%s' not found", title)
}

// GetSectionContent returns the content of a specific section
func (ds *DocumentStructure) GetSectionContent(sectionID string) ([]types.Paragraph, error) {
	for _, section := range ds.Sections {
		if section.ID == sectionID {
			return section.Paragraphs, nil
		}
	}
	return nil, fmt.Errorf("section with ID '%s' not found", sectionID)
}

// isHeadingParagraph checks if a paragraph is a heading
func isHeadingParagraph(paragraph types.Paragraph) bool {
	// 检查段落样式
	if strings.Contains(strings.ToLower(paragraph.Style), "heading") {
		return true
	}
	
	// 检查是否有粗体格式
	for _, run := range paragraph.Runs {
		if run.Bold && len(strings.TrimSpace(run.Text)) > 0 {
			return true
		}
	}
	
	return false
}

// extractHeadingTitle extracts the title from a heading paragraph
func extractHeadingTitle(paragraph types.Paragraph) string {
	title := ""
	for _, run := range paragraph.Runs {
		title += run.Text
	}
	return strings.TrimSpace(title)
}

// determineHeadingLevel determines the heading level
func determineHeadingLevel(paragraph types.Paragraph) int {
	// 根据样式名称判断级别
	style := strings.ToLower(paragraph.Style)
	
	switch {
	case strings.Contains(style, "heading1") || strings.Contains(style, "title"):
		return 1
	case strings.Contains(style, "heading2"):
		return 2
	case strings.Contains(style, "heading3"):
		return 3
	case strings.Contains(style, "heading4"):
		return 4
	case strings.Contains(style, "heading5"):
		return 5
	case strings.Contains(style, "heading6"):
		return 6
	default:
		// 根据粗体程度判断
		boldCount := 0
		for _, run := range paragraph.Runs {
			if run.Bold {
				boldCount++
			}
		}
		
		if boldCount > 0 {
			return 1
		}
		return 1
	}
}

// SortSectionsByLevel sorts sections by their level
func (ds *DocumentStructure) SortSectionsByLevel() {
	sort.Slice(ds.Sections, func(i, j int) bool {
		if ds.Sections[i].Level != ds.Sections[j].Level {
			return ds.Sections[i].Level < ds.Sections[j].Level
		}
		return ds.Sections[i].StartIndex < ds.Sections[j].StartIndex
	})
}

// GetOutlineAsString returns the outline as a formatted string
func (ds *DocumentStructure) GetOutlineAsString() string {
	var result strings.Builder
	result.WriteString("文档大纲:\n")
	
	for _, section := range ds.Outline.Sections {
		ds.writeSectionToString(&result, section, 0)
	}
	
	return result.String()
}

// writeSectionToString writes a section to the string builder
func (ds *DocumentStructure) writeSectionToString(builder *strings.Builder, section DocumentSection, level int) {
	indent := strings.Repeat("  ", level)
	builder.WriteString(fmt.Sprintf("%s%s (级别: %d)\n", indent, section.Title, section.Level))
	
	for _, subSection := range section.SubSections {
		ds.writeSectionToString(builder, subSection, level+1)
	}
} 