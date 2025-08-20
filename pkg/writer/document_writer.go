// Package writer provides document writing and modification functionality
package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/opc"
	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/word"
)

// DocumentWriter provides functionality to modify and create Word documents
type DocumentWriter struct {
	container      *opc.Container
	document       *word.Document
	commentManager *word.CommentManager // 使用新的批注管理器
}

// NewDocumentWriter creates a new document writer
func NewDocumentWriter() *DocumentWriter {
	return &DocumentWriter{
		commentManager: word.NewCommentManager(),
	}
}

// OpenForModification opens an existing document for modification
func (w *DocumentWriter) OpenForModification(filename string) error {
	doc, err := word.Open(filename)
	if err != nil {
		return fmt.Errorf("failed to open document for modification: %w", err)
	}

	w.document = doc
	w.container = doc.GetContainer()
	return nil
}

// CreateNewDocument creates a new empty Word document
func (w *DocumentWriter) CreateNewDocument() error {
	// Create a new OPC container
	w.container = &opc.Container{}

	// Create basic document structure
	w.document = &word.Document{}

	// Initialize with empty content
	mainPart := &word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	w.document.SetMainPart(mainPart)

	return nil
}

// AddParagraph adds a new paragraph to the document
func (w *DocumentWriter) AddParagraph(text string, style string) error {
	if w.document == nil || w.document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	paragraph := types.Paragraph{
		Text:  text,
		Style: style,
		Runs: []types.Run{
			{
				Text: text,
			},
		},
	}

	mainPart := w.document.GetMainPart()
	mainPart.Content.Paragraphs = append(
		mainPart.Content.Paragraphs, paragraph)

	// Update document text
	mainPart.Content.Text += text + "\n"

	return nil
}

// AddFormattedParagraph adds a paragraph with specific formatting
func (w *DocumentWriter) AddFormattedParagraph(text string, style string, runs []types.Run) error {
	if w.document == nil || w.document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	paragraph := types.Paragraph{
		Text:  text,
		Style: style,
		Runs:  runs,
	}

	mainPart := w.document.GetMainPart()
	mainPart.Content.Paragraphs = append(
		mainPart.Content.Paragraphs, paragraph)

	// Update document text
	mainPart.Content.Text += text + "\n"

	return nil
}

// AddTable adds a new table to the document
func (w *DocumentWriter) AddTable(rows [][]string) error {
	if w.document == nil || w.document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	table := types.Table{
		Rows: make([]types.TableRow, len(rows)),
	}

	for i, rowData := range rows {
		row := types.TableRow{
			Cells: make([]types.TableCell, len(rowData)),
		}

		for j, cellText := range rowData {
			row.Cells[j] = types.TableCell{
				Text: cellText,
			}
		}

		table.Rows[i] = row
	}

	if len(rows) > 0 {
		table.Columns = len(rows[0])
	}

	mainPart := w.document.GetMainPart()
	mainPart.Content.Tables = append(
		mainPart.Content.Tables, table)

	return nil
}

// ReplaceText replaces all occurrences of old text with new text
func (w *DocumentWriter) ReplaceText(oldText, newText string) error {
	if w.document == nil || w.document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	mainPart := w.document.GetMainPart()

	// Replace in document text
	mainPart.Content.Text = strings.ReplaceAll(
		mainPart.Content.Text, oldText, newText)

	// Replace in paragraphs
	for i := range mainPart.Content.Paragraphs {
		paragraph := &mainPart.Content.Paragraphs[i]
		paragraph.Text = strings.ReplaceAll(paragraph.Text, oldText, newText)

		// Replace in runs
		for j := range paragraph.Runs {
			run := &paragraph.Runs[j]
			run.Text = strings.ReplaceAll(run.Text, oldText, newText)
		}
	}

	// Replace in table cells
	for i := range mainPart.Content.Tables {
		table := &mainPart.Content.Tables[i]
		for j := range table.Rows {
			row := &table.Rows[j]
			for k := range row.Cells {
				cell := &row.Cells[k]
				cell.Text = strings.ReplaceAll(cell.Text, oldText, newText)
			}
		}
	}

	return nil
}

// SetParagraphStyle sets the style of a specific paragraph
func (w *DocumentWriter) SetParagraphStyle(index int, style string) error {
	if w.document == nil || w.document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	mainPart := w.document.GetMainPart()
	if index < 0 || index >= len(mainPart.Content.Paragraphs) {
		return fmt.Errorf("paragraph index out of range")
	}

	mainPart.Content.Paragraphs[index].Style = style
	return nil
}

// SetRunFormatting sets formatting for a specific run in a paragraph
func (w *DocumentWriter) SetRunFormatting(paragraphIndex, runIndex int, formatting types.Run) error {
	if w.document == nil || w.document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	mainPart := w.document.GetMainPart()
	if paragraphIndex < 0 || paragraphIndex >= len(mainPart.Content.Paragraphs) {
		return fmt.Errorf("paragraph index out of range")
	}

	paragraph := &mainPart.Content.Paragraphs[paragraphIndex]
	if runIndex < 0 || runIndex >= len(paragraph.Runs) {
		return fmt.Errorf("run index out of range")
	}

	paragraph.Runs[runIndex] = formatting
	return nil
}

// AddComment adds a comment to the document
// Based on Open-XML-SDK AddComment method
func (w *DocumentWriter) AddComment(author, text, paragraphText string) error {
	if w.document == nil {
		return fmt.Errorf("document not initialized")
	}

	// 使用新的批注管理器添加批注
	comment, err := w.commentManager.AddComment(author, text, "para_1", "run_1", 0, len(paragraphText))
	if err != nil {
		return fmt.Errorf("failed to add comment: %w", err)
	}

	// 在文档中添加批注引用
	return w.addCommentReferenceToDocument(comment.ID, paragraphText)
}

// addCommentReferenceToDocument 在文档中添加批注引用
func (w *DocumentWriter) addCommentReferenceToDocument(commentID, paragraphText string) error {
	if w.document == nil || w.document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	mainPart := w.document.GetMainPart()

	// 查找匹配的段落并添加批注引用
	for i, paragraph := range mainPart.Content.Paragraphs {
		if strings.Contains(paragraph.Text, paragraphText) {
			// 在段落中添加批注范围标记和引用
			mainPart.Content.Paragraphs[i].HasComment = true
			mainPart.Content.Paragraphs[i].CommentID = commentID
			return nil
		}
	}

	return fmt.Errorf("未找到匹配的段落文本: %s", paragraphText)
}

// Save saves the document to a file
func (w *DocumentWriter) Save(filename string) error {
	if w.document == nil {
		return fmt.Errorf("document not initialized")
	}

	// Generate XML content for the main document part
	xmlContent, err := w.generateDocumentXML()
	if err != nil {
		return fmt.Errorf("failed to generate document XML: %w", err)
	}

	// Create a new container with the modified content
	container := &opc.Container{}

	// Add the main document part
	container.AddPart(
		"word/document.xml",
		xmlContent,
		"application/vnd.openxmlformats-officedocument.word.document.main+xml",
	)

	// Add comments part if there are comments
	if len(w.commentManager.Comments) > 0 {
		commentsXML := w.generateCommentsXML()
		container.AddPart(
			"word/comments.xml",
			commentsXML,
			"application/vnd.openxmlformats-officedocument.word.comments+xml",
		)
	}

	// Add [Content_Types].xml
	contentTypesXML := w.generateContentTypesXML()
	container.AddPart(
		"[Content_Types].xml",
		contentTypesXML,
		"application/vnd.openxmlformats-package.content-types",
	)

	// Add _rels/.rels
	rootRelsXML := w.generateRootRelsXML()
	container.AddPart(
		"_rels/.rels",
		rootRelsXML,
		"application/vnd.openxmlformats-package.relationships+xml",
	)

	// Add word/_rels/document.xml.rels
	documentRelsXML := w.generateDocumentRelsXML()
	container.AddPart(
		"word/_rels/document.xml.rels",
		documentRelsXML,
		"application/vnd.openxmlformats-package.relationships+xml",
	)

	// Add settings.xml if there are comments (for WPS compatibility)
	if len(w.commentManager.Comments) > 0 {
		settingsXML := w.generateSettingsXML()
		container.AddPart(
			"word/settings.xml",
			settingsXML,
			"application/vnd.openxmlformats-officedocument.word.settings+xml",
		)
	}

	// Save the container to file
	return container.SaveToFile(filename)
}

// generateCommentsXML generates the comments XML content
func (w *DocumentWriter) generateCommentsXML() []byte {
	return []byte(w.commentManager.GenerateCommentsXML())
}

// generateDocumentXML generates the XML content for the main document part
func (w *DocumentWriter) generateDocumentXML() ([]byte, error) {
	if w.document == nil || w.document.GetMainPart() == nil {
		return nil, fmt.Errorf("document not initialized")
	}

	mainPart := w.document.GetMainPart()

	// Create the XML structure
	doc := &DocumentXML{
		XMLName: xml.Name{Local: "w:document"},
		XMLNS:   "http://schemas.openxmlformats.org/word/2006/main",
		Body: DocumentBody{
			XMLName: xml.Name{Local: "w:body"},
		},
	}

	// Add paragraphs
	for _, paragraph := range mainPart.Content.Paragraphs {
		xmlParagraph := ParagraphXML{
			XMLName: xml.Name{Local: "w:p"},
		}

		// Add paragraph properties if style is set
		if paragraph.Style != "" {
			xmlParagraph.Properties = &ParagraphPropertiesXML{
				XMLName: xml.Name{Local: "w:pPr"},
				Style: &StyleXML{
					XMLName: xml.Name{Local: "w:pStyle"},
					Val:     paragraph.Style,
				},
			}
		}

		// Add comment range start if paragraph has comment
		if paragraph.HasComment {
			xmlParagraph.CommentRangeStart = &CommentRangeStartXML{
				ID: paragraph.CommentID,
			}
		}

		// Add runs
		for _, run := range paragraph.Runs {
			xmlRun := RunXML{
				XMLName: xml.Name{Local: "w:r"},
			}

			// Add run properties if formatting is set
			if run.Bold || run.Italic || run.Underline || run.FontSize > 0 || run.FontName != "" {
				xmlRun.Properties = &RunPropertiesXML{
					XMLName: xml.Name{Local: "w:rPr"},
				}

				if run.Bold {
					xmlRun.Properties.Bold = &BoldXML{
						XMLName: xml.Name{Local: "w:b"},
						Val:     "true",
					}
				}

				if run.Italic {
					xmlRun.Properties.Italic = &ItalicXML{
						XMLName: xml.Name{Local: "w:i"},
						Val:     "true",
					}
				}

				if run.Underline {
					xmlRun.Properties.Underline = &UnderlineXML{
						XMLName: xml.Name{Local: "w:u"},
						Val:     "single",
					}
				}

				if run.FontSize > 0 {
					xmlRun.Properties.Size = &SizeXML{
						XMLName: xml.Name{Local: "w:sz"},
						Val:     fmt.Sprintf("%d", run.FontSize),
					}
				}

				if run.FontName != "" {
					xmlRun.Properties.Font = &FontXML{
						XMLName: xml.Name{Local: "w:rFonts"},
						Ascii:   run.FontName,
						HAnsi:   run.FontName,
					}
				}
			}

			// Add text
			xmlRun.Text = &TextXML{
				XMLName: xml.Name{Local: "w:t"},
				Content: run.Text,
			}

			xmlParagraph.Runs = append(xmlParagraph.Runs, xmlRun)
		}

		// Add comment range end and reference if paragraph has comment
		if paragraph.HasComment {
			xmlParagraph.CommentRangeEnd = &CommentRangeEndXML{
				ID: paragraph.CommentID,
			}
			xmlParagraph.CommentReference = &CommentReferenceXML{
				ID: paragraph.CommentID,
			}
		}

		doc.Body.Paragraphs = append(doc.Body.Paragraphs, xmlParagraph)
	}

	// Add tables
	for _, table := range mainPart.Content.Tables {
		xmlTable := TableXML{
			XMLName: xml.Name{Local: "w:tbl"},
		}

		for _, row := range table.Rows {
			xmlRow := TableRowXML{
				XMLName: xml.Name{Local: "w:tr"},
			}

			for _, cell := range row.Cells {
				xmlCell := TableCellXML{
					XMLName: xml.Name{Local: "w:tc"},
					Content: []interface{}{
						ParagraphXML{
							XMLName: xml.Name{Local: "w:p"},
							Runs: []RunXML{
								{
									XMLName: xml.Name{Local: "w:r"},
									Text: &TextXML{
										XMLName: xml.Name{Local: "w:t"},
										Content: cell.Text,
									},
								},
							},
						},
					},
				}
				xmlRow.Cells = append(xmlRow.Cells, xmlCell)
			}

			xmlTable.Rows = append(xmlTable.Rows, xmlRow)
		}

		doc.Body.Tables = append(doc.Body.Tables, xmlTable)
	}

	// Marshal to XML
	var buf bytes.Buffer
	buf.WriteString(xml.Header)
	encoder := xml.NewEncoder(&buf)
	encoder.Indent("", "  ")
	if err := encoder.Encode(doc); err != nil {
		return nil, fmt.Errorf("failed to encode document XML: %w", err)
	}

	return buf.Bytes(), nil
}

// generateContentTypesXML generates the XML content for [Content_Types].xml
func (w *DocumentWriter) generateContentTypesXML() []byte {
	contentTypesXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Default Extension="tiff" ContentType="image/tiff"/>
  <Default Extension="bmp" ContentType="image/bmp"/>
  <Default Extension="wmf" ContentType="image/wmf"/>
  <Default Extension="emf" ContentType="image/emf"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.word.document.main+xml"/>`

	// Add comments content type if there are comments
	if len(w.commentManager.Comments) > 0 {
		contentTypesXML += `
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.word.comments+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.word.settings+xml"/>`
	}

	contentTypesXML += `
</Types>`
	return []byte(contentTypesXML)
}

// generateRootRelsXML generates the XML content for _rels/.rels
func (w *DocumentWriter) generateRootRelsXML() []byte {
	rootRelsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

	return []byte(rootRelsXML)
}

// generateDocumentRelsXML generates the XML content for word/_rels/document.xml.rels
func (w *DocumentWriter) generateDocumentRelsXML() []byte {
	documentRelsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`

	// Add comments relationship if there are comments
	if len(w.commentManager.Comments) > 0 {
		documentRelsXML += `
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>`
	}

	documentRelsXML += `
</Relationships>`

	return []byte(documentRelsXML)
}

// generateSettingsXML generates the XML content for word/settings.xml
func (w *DocumentWriter) generateSettingsXML() []byte {
	settingsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/word/2006/main">
  <w:showComments w:val="true"/>
  <w:trackRevisions w:val="false"/>
  <w:printComments w:val="true"/>
  <w:printHiddenText w:val="false"/>
  <w:printBackground w:val="false"/>
  <w:zoom w:percent="100"/>
</w:settings>`
	return []byte(settingsXML)
}

// XML structures for document generation
type DocumentXML struct {
	XMLName xml.Name `xml:"w:document"`
	XMLNS   string   `xml:"xmlns:w,attr"`
	Body    DocumentBody
}

type DocumentBody struct {
	XMLName    xml.Name       `xml:"w:body"`
	Paragraphs []ParagraphXML `xml:"w:p"`
	Tables     []TableXML     `xml:"w:tbl"`
}

type ParagraphXML struct {
	XMLName           xml.Name                `xml:"w:p"`
	Properties        *ParagraphPropertiesXML `xml:"w:pPr,omitempty"`
	CommentRangeStart *CommentRangeStartXML   `xml:"w:commentRangeStart,omitempty"`
	Runs              []RunXML                `xml:"w:r"`
	CommentRangeEnd   *CommentRangeEndXML     `xml:"w:commentRangeEnd,omitempty"`
	CommentReference  *CommentReferenceXML    `xml:"w:commentReference,omitempty"`
}

type ParagraphPropertiesXML struct {
	XMLName xml.Name  `xml:"w:pPr"`
	Style   *StyleXML `xml:"w:pStyle,omitempty"`
}

type StyleXML struct {
	XMLName xml.Name `xml:"w:pStyle"`
	Val     string   `xml:"w:val,attr"`
}

type RunXML struct {
	XMLName    xml.Name          `xml:"w:r"`
	Properties *RunPropertiesXML `xml:"w:rPr,omitempty"`
	Text       *TextXML          `xml:"w:t,omitempty"`
}

type RunPropertiesXML struct {
	XMLName   xml.Name      `xml:"w:rPr"`
	Bold      *BoldXML      `xml:"w:b,omitempty"`
	Italic    *ItalicXML    `xml:"w:i,omitempty"`
	Underline *UnderlineXML `xml:"w:u,omitempty"`
	Size      *SizeXML      `xml:"w:sz,omitempty"`
	Font      *FontXML      `xml:"w:rFonts,omitempty"`
}

type TextXML struct {
	XMLName xml.Name `xml:"w:t"`
	Content string   `xml:",chardata"`
}

type BoldXML struct {
	XMLName xml.Name `xml:"w:b"`
	Val     string   `xml:"w:val,attr"`
}

type ItalicXML struct {
	XMLName xml.Name `xml:"w:i"`
	Val     string   `xml:"w:val,attr"`
}

type UnderlineXML struct {
	XMLName xml.Name `xml:"w:u"`
	Val     string   `xml:"w:val,attr"`
}

type SizeXML struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr"`
}

type FontXML struct {
	XMLName xml.Name `xml:"w:rFonts"`
	Ascii   string   `xml:"w:ascii,attr"`
	HAnsi   string   `xml:"w:hAnsi,attr"`
}

// CommentRangeStartXML represents the start of a comment range
type CommentRangeStartXML struct {
	ID string `xml:"w:id,attr"`
}

// CommentRangeEndXML represents the end of a comment range
type CommentRangeEndXML struct {
	ID string `xml:"w:id,attr"`
}

// CommentReferenceXML represents a comment reference
type CommentReferenceXML struct {
	ID string `xml:"w:id,attr"`
}

type TableXML struct {
	XMLName xml.Name      `xml:"w:tbl"`
	Rows    []TableRowXML `xml:"w:tr"`
}

type TableRowXML struct {
	XMLName xml.Name       `xml:"w:tr"`
	Cells   []TableCellXML `xml:"w:tc"`
}

type TableCellXML struct {
	XMLName xml.Name      `xml:"w:tc"`
	Content []interface{} `xml:",any"`
}
