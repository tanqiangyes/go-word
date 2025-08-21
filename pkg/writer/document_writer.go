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
	Container      *opc.Container
	Document       *word.Document
	CommentManager *word.CommentManager // 使用新的批注管理器
}

// NewDocumentWriter creates a new document writer
func NewDocumentWriter() *DocumentWriter {
	return &DocumentWriter{
		CommentManager: word.NewCommentManager(),
	}
}

// OpenForModification opens an existing document for modification
func (w *DocumentWriter) OpenForModification(filename string) error {
	doc, err := word.Open(filename)
	if err != nil {
		return fmt.Errorf("failed to open document for modification: %w", err)
	}

	w.Document = doc
	w.Container = doc.GetContainer()
	return nil
}

// CreateNewDocument creates a new empty Word document
func (w *DocumentWriter) CreateNewDocument() error {
	// Create a new OPC container
	w.Container = &opc.Container{}

	// Create basic document structure
	w.Document = &word.Document{}

	// Initialize with empty content
	mainPart := &word.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
			Text:       "",
		},
	}
	w.Document.SetMainPart(mainPart)

	return nil
}

// AddParagraph adds a new paragraph to the document
func (w *DocumentWriter) AddParagraph(text string, style string) error {
	if w.Document == nil || w.Document.GetMainPart() == nil {
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

	mainPart := w.Document.GetMainPart()
	mainPart.Content.Paragraphs = append(
		mainPart.Content.Paragraphs, paragraph)

	// Update document text
	mainPart.Content.Text += text + "\n"

	return nil
}

// AddFormattedParagraph adds a paragraph with specific formatting
func (w *DocumentWriter) AddFormattedParagraph(text string, style string, runs []types.Run) error {
	if w.Document == nil || w.Document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	paragraph := types.Paragraph{
		Text:  text,
		Style: style,
		Runs:  runs,
	}

	mainPart := w.Document.GetMainPart()
	mainPart.Content.Paragraphs = append(
		mainPart.Content.Paragraphs, paragraph)

	// Update document text
	mainPart.Content.Text += text + "\n"

	return nil
}

// AddTable adds a new table to the document
func (w *DocumentWriter) AddTable(rows [][]string) error {
	if w.Document == nil || w.Document.GetMainPart() == nil {
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

	mainPart := w.Document.GetMainPart()
	mainPart.Content.Tables = append(
		mainPart.Content.Tables, table)

	return nil
}

// ReplaceText replaces all occurrences of old text with new text
func (w *DocumentWriter) ReplaceText(oldText, newText string) error {
	if w.Document == nil || w.Document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	mainPart := w.Document.GetMainPart()

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
	if w.Document == nil || w.Document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	mainPart := w.Document.GetMainPart()
	if index < 0 || index >= len(mainPart.Content.Paragraphs) {
		return fmt.Errorf("paragraph index out of range")
	}

	mainPart.Content.Paragraphs[index].Style = style
	return nil
}

// SetRunFormatting sets formatting for a specific run in a paragraph
func (w *DocumentWriter) SetRunFormatting(paragraphIndex, runIndex int, formatting types.Run) error {
	if w.Document == nil || w.Document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	mainPart := w.Document.GetMainPart()
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
	if w.Document == nil {
		return fmt.Errorf("document not initialized")
	}

	// 使用新的批注管理器添加批注
	comment, err := w.CommentManager.AddComment(author, text, "para_1", "run_1", 0, len(paragraphText))
	if err != nil {
		return fmt.Errorf("failed to add comment: %w", err)
	}

	// 在文档中添加批注引用
	return w.addCommentReferenceToDocument(comment.ID, paragraphText)
}

// addCommentReferenceToDocument 在文档中添加批注引用
func (w *DocumentWriter) addCommentReferenceToDocument(commentID, paragraphText string) error {
	if w.Document == nil || w.Document.GetMainPart() == nil {
		return fmt.Errorf("document not initialized")
	}

	mainPart := w.Document.GetMainPart()

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
	if w.Document == nil {
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
		"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
	)

	// Add comments part if there are comments
	if len(w.CommentManager.Comments) > 0 {
		commentsXML := w.generateCommentsXML()
		container.AddPart(
			"word/comments.xml",
			commentsXML,
			"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
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

	// Add styles.xml (required for proper document display)
	stylesXML := w.generateStylesXML()
	container.AddPart(
		"word/styles.xml",
		stylesXML,
		"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
	)

	// Add settings.xml (always required for compatibility)
	settingsXML := w.generateSettingsXML()
	container.AddPart(
		"word/settings.xml",
		settingsXML,
		"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
	)

	// Add fontTable.xml (required for font support)
	fontTableXML := w.generateFontTableXML()
	container.AddPart(
		"word/fontTable.xml",
		fontTableXML,
		"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
	)

	// Add theme/theme1.xml (required for theme support)
	themeXML := w.generateThemeXML()
	container.AddPart(
		"word/theme/theme1.xml",
		themeXML,
		"application/vnd.openxmlformats-officedocument.theme+xml",
	)

	// Add docProps/app.xml (required for application properties)
	appXML := w.generateAppXML()
	container.AddPart(
		"docProps/app.xml",
		appXML,
		"application/vnd.openxmlformats-officedocument.extended-properties+xml",
	)

	// Add docProps/core.xml (required for core properties)
	coreXML := w.generateCoreXML()
	container.AddPart(
		"docProps/core.xml",
		coreXML,
		"application/vnd.openxmlformats-package.core-properties+xml",
	)

	// Add docProps/custom.xml (required for custom properties)
	customXML := w.generateCustomXML()
	container.AddPart(
		"docProps/custom.xml",
		customXML,
		"application/vnd.openxmlformats-officedocument.custom-properties+xml",
	)

	// Save the container to file
	return container.SaveToFile(filename)
}

// generateCommentsXML generates the comments XML content
func (w *DocumentWriter) generateCommentsXML() []byte {
	commentsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">`

	for _, comment := range w.CommentManager.Comments {
		commentsXML += fmt.Sprintf(`
  <w:comment w:id="%s" w:author="%s" w:date="%s" w:initials="%s" w:paraId="%s" w:time="%s" w:index="0">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="CommentText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
          <w:sz w:val="20"/>
          <w:szCs w:val="20"/>
        </w:rPr>
        <w:t xml:space="preserve">%s</w:t>
      </w:r>
    </w:p>
  </w:comment>`, comment.ID, comment.Author, comment.Date, comment.Initials, comment.ID, comment.Date, comment.Text)
	}

	commentsXML += `
</w:comments>`
	return []byte(commentsXML)
}

// generateDocumentXML generates the XML content for the main document part
func (w *DocumentWriter) generateDocumentXML() ([]byte, error) {
	if w.Document == nil || w.Document.GetMainPart() == nil {
		return nil, fmt.Errorf("document not initialized")
	}

	mainPart := w.Document.GetMainPart()

	// Create the XML structure
	doc := &DocumentXML{
		XMLName: xml.Name{Local: "w:document"},
		XMLNS:   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XMLNSMC: "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XMLNSR: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XMLNSW14: "http://schemas.microsoft.com/office/word/2010/wordml",
		MCIgnorable: "w14",
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
		} else {
			// Always add paragraph properties for compatibility
			xmlParagraph.Properties = &ParagraphPropertiesXML{
				XMLName: xml.Name{Local: "w:pPr"},
				Style: &StyleXML{
					XMLName: xml.Name{Local: "w:pStyle"},
					Val:     "Normal",
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

			// Add run properties only if there's actual formatting
			hasFormatting := run.Bold || run.Italic || run.Underline || run.FontSize > 0 || run.FontName != ""
			if hasFormatting {
				xmlRun.Properties = &RunPropertiesXML{
					XMLName: xml.Name{Local: "w:rPr"},
				}

				// Add specific formatting if set
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

		// Add table properties for better compatibility
		xmlTable.Properties = &TablePropertiesXML{
			XMLName: xml.Name{Local: "w:tblPr"},
			TableBorders: &TableBordersXML{
				XMLName: xml.Name{Local: "w:tblBorders"},
				Top:     &TopBorderXML{XMLName: xml.Name{Local: "w:top"}, Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Left:    &LeftBorderXML{XMLName: xml.Name{Local: "w:left"}, Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Bottom:  &BottomBorderXML{XMLName: xml.Name{Local: "w:bottom"}, Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Right:   &RightBorderXML{XMLName: xml.Name{Local: "w:right"}, Val: "single", Sz: "4", Space: "0", Color: "auto"},
				InsideH: &InsideHBorderXML{XMLName: xml.Name{Local: "w:insideH"}, Val: "single", Sz: "4", Space: "0", Color: "auto"},
				InsideV: &InsideVBorderXML{XMLName: xml.Name{Local: "w:insideV"}, Val: "single", Sz: "4", Space: "0", Color: "auto"},
			},
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

	// Add section properties for page settings
	doc.Body.SectionProperties = &SectionPropertiesXML{
		XMLName: xml.Name{Local: "w:sectPr"},
		PageSize: &PageSizeXML{
			XMLName: xml.Name{Local: "w:pgSz"},
			Width:   "11906",
			Height:  "16838",
		},
		PageMargins: &PageMarginsXML{
			XMLName: xml.Name{Local: "w:pgMar"},
			Top:     "1440",
			Right:   "1800",
			Bottom:  "1440",
			Left:    "1800",
			Header:  "851",
			Footer:  "992",
			Gutter:  "0",
		},
		Columns: &ColumnsXML{
			XMLName: xml.Name{Local: "w:cols"},
			Space:   "425",
			Number:  "1",
		},
		DocumentGrid: &DocumentGridXML{
			XMLName: xml.Name{Local: "w:docGrid"},
			Type:    "lines",
			LinePitch: "312",
			CharSpace: "0",
		},
	}

	// Marshal to XML
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
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
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`

	// Add all required content types for compatibility
	contentTypesXML += `
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`

	// Add comments content type if there are comments
	if len(w.CommentManager.Comments) > 0 {
		contentTypesXML += `
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

	return []byte(rootRelsXML)
}

// generateDocumentRelsXML generates the XML content for word/_rels/document.xml.rels
func (w *DocumentWriter) generateDocumentRelsXML() []byte {
	documentRelsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`

	// Add comments relationship if there are comments
	relationshipId := 5
	if len(w.CommentManager.Comments) > 0 {
		documentRelsXML += fmt.Sprintf(`
  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>`, relationshipId)
	}

	documentRelsXML += `
</Relationships>`

	return []byte(documentRelsXML)
}

// generateSettingsXML generates the XML content for word/settings.xml
func (w *DocumentWriter) generateSettingsXML() []byte {
	settingsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">
  <w:showComments w:val="true"/>
  <w:trackRevisions w:val="false"/>
  <w:printComments w:val="true"/>
  <w:printHiddenText w:val="false"/>
  <w:printBackground w:val="false"/>
  <w:zoom w:percent="100"/>
</w:settings>`
	return []byte(settingsXML)
}

// generateStylesXML generates the XML content for word/styles.xml
func (w *DocumentWriter) generateStylesXML() []byte {
	stylesXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Times New Roman"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:widowControl w:val="0"/>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
  <w:style w:type="numbering" w:default="1" w:styleId="NoList">
    <w:name w:val="No List"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CommentText">
    <w:name w:val="Comment Text"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="CommentTextChar"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="CommentTextChar">
    <w:name w:val="Comment Text Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="CommentText"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BalloonText">
    <w:name w:val="Balloon Text"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="BalloonTextChar"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="BalloonTextChar">
    <w:name w:val="Balloon Text Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="BalloonText"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>
</w:styles>`
	return []byte(stylesXML)
}

// XML structures for document generation
type DocumentXML struct {
	XMLName xml.Name `xml:"w:document"`
	XMLNS   string   `xml:"xmlns:w,attr"`
	XMLNSMC string `xml:"xmlns:mc,attr"`
	XMLNSR string `xml:"xmlns:r,attr"`
	XMLNSW14 string `xml:"xmlns:w14,attr"`
	MCIgnorable string `xml:"mc:Ignorable,attr"`
	Body    DocumentBody
}

type DocumentBody struct {
	XMLName    xml.Name       `xml:"w:body"`
	Paragraphs []ParagraphXML `xml:"w:p"`
	Tables     []TableXML     `xml:"w:tbl"`
	SectionProperties *SectionPropertiesXML `xml:"w:sectPr,omitempty"`
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
	XMLName xml.Name `xml:"w:commentRangeStart"`
	ID      string   `xml:"w:id,attr"`
}

// CommentRangeEndXML represents the end of a comment range
type CommentRangeEndXML struct {
	XMLName xml.Name `xml:"w:commentRangeEnd"`
	ID      string   `xml:"w:id,attr"`
}

// CommentReferenceXML represents a comment reference
type CommentReferenceXML struct {
	XMLName xml.Name `xml:"w:commentReference"`
	ID      string   `xml:"w:id,attr"`
}

type TableXML struct {
	XMLName xml.Name      `xml:"w:tbl"`
	Properties *TablePropertiesXML `xml:"w:tblPr,omitempty"`
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

// generateFontTableXML generates the XML content for word/fontTable.xml
func (w *DocumentWriter) generateFontTableXML() []byte {
	fontTableXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="00"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="20007A87" w:usb1="80000000" w:usb2="00000008" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="宋体">
    <w:panose1 w:val="02010600030101010101"/>
    <w:charset w:val="86"/>
    <w:family w:val="auto"/>
    <w:pitch w:val="default"/>
    <w:sig w:usb0="00000203" w:usb1="288F0000" w:usb2="00000006" w:usb3="00000000" w:csb0="00040001" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Wingdings">
    <w:panose1 w:val="05000000000000000000"/>
    <w:charset w:val="02"/>
    <w:family w:val="auto"/>
    <w:pitch w:val="default"/>
    <w:sig w:usb0="00000000" w:usb1="00000000" w:usb2="00000000" w:usb3="00000000" w:csb0="80000000" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Arial">
    <w:panose1 w:val="020B0604020202020204"/>
    <w:charset w:val="01"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="default"/>
    <w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="400001FF" w:csb1="FFFF0000"/>
  </w:font>
  <w:font w:name="黑体">
    <w:panose1 w:val="02010609060101010101"/>
    <w:charset w:val="86"/>
    <w:family w:val="auto"/>
    <w:pitch w:val="default"/>
    <w:sig w:usb0="800002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="00040001" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Courier New">
    <w:panose1 w:val="02070309020205020404"/>
    <w:charset w:val="01"/>
    <w:family w:val="modern"/>
    <w:pitch w:val="default"/>
    <w:sig w:usb0="E0002EFF" w:usb1="C0007843" w:usb2="00000009" w:usb3="00000000" w:csb0="400001FF" w:csb1="FFFF0000"/>
  </w:font>
  <w:font w:name="Symbol">
    <w:panose1 w:val="05050102010706020507"/>
    <w:charset w:val="02"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="default"/>
    <w:sig w:usb0="00000000" w:usb1="00000000" w:usb2="00000000" w:usb3="00000000" w:csb0="80000000" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="default"/>
    <w:sig w:usb0="E4002EFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="200001FF" w:csb1="00000000"/>
  </w:font>
</w:fonts>`
	return []byte(fontTableXML)
}

// generateThemeXML generates the XML content for word/theme/theme1.xml
func (w *DocumentWriter) generateThemeXML() []byte {
	themeXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000"/>
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF"/>
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="1F497D"/>
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="EEECE1"/>
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="4F81BD"/>
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="C0504D"/>
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="9BBB59"/>
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="8064A2"/>
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="4BACC6"/>
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="F79646"/>
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0000FF"/>
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="800080"/>
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="35000">
              <a:schemeClr val="phClr">
                <a:tint val="37000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:shade val="51000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="80000">
              <a:schemeClr val="phClr">
                <a:shade val="93000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="94000"/>
                <a:satMod val="135000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
          <a:headEnd/>
          <a:tailEnd/>
        </a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
          <a:headEnd/>
          <a:tailEnd/>
        </a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
          <a:headEnd/>
          <a:tailEnd/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blur="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="63000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:solidFill>
          <a:schemeClr val="phClr">
            <a:tint val="15000"/>
            <a:satMod val="350000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="40000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="40000">
              <a:schemeClr val="phClr">
                <a:tint val="45000"/>
                <a:satMod val="350000"/>
                <a:shade val="99000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="20000"/>
                <a:satMod val="255000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
          </a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="80000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="30000"/>
                <a:satMod val="200000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
          </a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:themeElements>
    <a:objectDefaults/>
    <a:extraClrSchemeLst/>
  </a:theme>`
	return []byte(themeXML)
}

// generateAppXML generates the XML content for docProps/app.xml
func (w *DocumentWriter) generateAppXML() []byte {
	appXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <LinksUpToDate>false</LinksUpToDate>
  <CharactersWithSpaces>0</CharactersWithSpaces>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`
	return []byte(appXML)
}

// generateCoreXML generates the XML content for docProps/core.xml
func (w *DocumentWriter) generateCoreXML() []byte {
	coreXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Go-Word Generated Document</dc:title>
  <dc:subject>Document generated by Go-Word library</dc:subject>
  <dc:creator>Go-Word Library</dc:creator>
  <cp:keywords>DOCX, Word, Go, Library</cp:keywords>
  <dc:description>This document was generated using the Go-Word library</dc:description>
  <cp:lastModifiedBy>Go-Word Library</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`
	return []byte(coreXML)
}

// generateCustomXML generates the XML content for docProps/custom.xml
func (w *DocumentWriter) generateCustomXML() []byte {
	customXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Custom>
    <Application>Go-Word</Application>
    <DocSecurity>0</DocSecurity>
    <ScaleCrop>false</ScaleCrop>
    <LinksUpToDate>false</LinksUpToDate>
    <CharactersWithSpaces>0</CharactersWithSpaces>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>16.0000</AppVersion>
  </Custom>
</Properties>`
	return []byte(customXML)
}

// TablePropertiesXML represents the properties of a table
type TablePropertiesXML struct {
	XMLName      xml.Name         `xml:"w:tblPr"`
	TableBorders *TableBordersXML `xml:"w:tblBorders,omitempty"`
}

// TableBordersXML represents the borders of a table
type TableBordersXML struct {
	XMLName xml.Name         `xml:"w:tblBorders"`
	Top     *TopBorderXML    `xml:"w:top,omitempty"`
	Left    *LeftBorderXML   `xml:"w:left,omitempty"`
	Bottom  *BottomBorderXML `xml:"w:bottom,omitempty"`
	Right   *RightBorderXML  `xml:"w:right,omitempty"`
	InsideH *InsideHBorderXML `xml:"w:insideH,omitempty"`
	InsideV *InsideVBorderXML `xml:"w:insideV,omitempty"`
}

// BorderXML represents a single border
type BorderXML struct {
	XMLName xml.Name `xml:"w:top"` // This is a placeholder, actual border types are more complex
	Val     string   `xml:"w:val,attr"`
	Sz      string   `xml:"w:sz,attr"`
	Space   string   `xml:"w:space,attr"`
	Color   string   `xml:"w:color,attr"`
}

// TopBorderXML represents the top border of a table
type TopBorderXML struct {
	XMLName xml.Name `xml:"w:top"`
	Val     string   `xml:"w:val,attr"`
	Sz      string   `xml:"w:sz,attr"`
	Space   string   `xml:"w:space,attr"`
	Color   string   `xml:"w:color,attr"`
}

// LeftBorderXML represents the left border of a table
type LeftBorderXML struct {
	XMLName xml.Name `xml:"w:left"`
	Val     string   `xml:"w:val,attr"`
	Sz      string   `xml:"w:sz,attr"`
	Space   string   `xml:"w:space,attr"`
	Color   string   `xml:"w:color,attr"`
}

// BottomBorderXML represents the bottom border of a table
type BottomBorderXML struct {
	XMLName xml.Name `xml:"w:bottom"`
	Val     string   `xml:"w:val,attr"`
	Sz      string   `xml:"w:sz,attr"`
	Space   string   `xml:"w:space,attr"`
	Color   string   `xml:"w:color,attr"`
}

// RightBorderXML represents the right border of a table
type RightBorderXML struct {
	XMLName xml.Name `xml:"w:right"`
	Val     string   `xml:"w:val,attr"`
	Sz      string   `xml:"w:sz,attr"`
	Space   string   `xml:"w:space,attr"`
	Color   string   `xml:"w:color,attr"`
}

// InsideHBorderXML represents the inside horizontal border of a table
type InsideHBorderXML struct {
	XMLName xml.Name `xml:"w:insideH"`
	Val     string   `xml:"w:val,attr"`
	Sz      string   `xml:"w:sz,attr"`
	Space   string   `xml:"w:space,attr"`
	Color   string   `xml:"w:color,attr"`
}

// InsideVBorderXML represents the inside vertical border of a table
type InsideVBorderXML struct {
	XMLName xml.Name `xml:"w:insideV"`
	Val     string   `xml:"w:val,attr"`
	Sz      string   `xml:"w:sz,attr"`
	Space   string   `xml:"w:space,attr"`
	Color   string   `xml:"w:color,attr"`
}

// SectionPropertiesXML represents the properties of a document section
type SectionPropertiesXML struct {
	XMLName xml.Name `xml:"w:sectPr"`
	PageSize *PageSizeXML `xml:"w:pgSz,omitempty"`
	PageMargins *PageMarginsXML `xml:"w:pgMar,omitempty"`
	Columns *ColumnsXML `xml:"w:cols,omitempty"`
	DocumentGrid *DocumentGridXML `xml:"w:docGrid,omitempty"`
}

// PageSizeXML represents the page size of a document section
type PageSizeXML struct {
	XMLName xml.Name `xml:"w:pgSz"`
	Width string `xml:"w:w,attr"`
	Height string `xml:"w:h,attr"`
}

// PageMarginsXML represents the page margins of a document section
type PageMarginsXML struct {
	XMLName xml.Name `xml:"w:pgMar"`
	Top string `xml:"w:top,attr"`
	Right string `xml:"w:right,attr"`
	Bottom string `xml:"w:bottom,attr"`
	Left string `xml:"w:left,attr"`
	Header string `xml:"w:header,attr"`
	Footer string `xml:"w:footer,attr"`
	Gutter string `xml:"w:gutter,attr"`
}

// ColumnsXML represents the column settings of a document section
type ColumnsXML struct {
	XMLName xml.Name `xml:"w:cols"`
	Space string `xml:"w:space,attr"`
	Number string `xml:"w:num,attr"`
}

// DocumentGridXML represents the document grid settings of a document section
type DocumentGridXML struct {
	XMLName xml.Name `xml:"w:docGrid"`
	Type string `xml:"w:type,attr"`
	LinePitch string `xml:"w:linePitch,attr"`
	CharSpace string `xml:"w:charSpace,attr"`
}
