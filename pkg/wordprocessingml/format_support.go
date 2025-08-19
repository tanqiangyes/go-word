// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/tanqiangyes/go-word/pkg/utils"
	"github.com/tanqiangyes/go-word/pkg/types"
)

// FormatSupport represents format support functionality
type FormatSupport struct {
	Document *Document
}

// RichTextContent represents rich text content
type RichTextContent struct {
	Text        string
	Formatting  RichTextFormatting
	Hyperlinks  []Hyperlink
	Images      []Image
	Tables      []RichTextTable
	Lists       []RichTextList
}

// RichTextFormatting represents rich text formatting
type RichTextFormatting struct {
	Font        Font
	Paragraph   ParagraphFormat
	Borders     BorderFormat
	Shading     ShadingFormat
	Effects     TextEffects
}

// Font represents font properties
type Font struct {
	Name      string
	Size      float64
	Bold      bool
	Italic    bool
	Underline bool
	Strike    bool
	Color     string
	Highlight string
	Subscript bool
	Superscript bool
	SmallCaps bool
	AllCaps   bool
}

// ParagraphFormat represents paragraph formatting
type ParagraphFormat struct {
	Alignment      string
	Indent         IndentFormat
	Spacing        SpacingFormat
	Borders        BorderFormat
	Shading        ShadingFormat
	KeepLines      bool
	KeepNext       bool
	PageBreakBefore bool
	WidowControl   bool
}

// IndentFormat represents indentation formatting
type IndentFormat struct {
	Left    float64
	Right   float64
	First   float64
	Hanging float64
}

// SpacingFormat represents spacing formatting
type SpacingFormat struct {
	Before  float64
	After   float64
	Line    float64
	Between bool
}

// BorderFormat represents border formatting
type BorderFormat struct {
	Top     BorderSide
	Bottom  BorderSide
	Left    BorderSide
	Right   BorderSide
	Between BorderSide
	Bar     BorderSide
}

// ShadingFormat represents shading formatting
type ShadingFormat struct {
	Fill      string
	Color     string
	ThemeFill string
	ThemeColor string
	Val       string
}

// TextEffects represents text effects
type TextEffects struct {
	Shadow     bool
	Outline    bool
	Emboss     bool
	Engrave    bool
	Reflection bool
	Glow       bool
	SoftEdge   bool
}

// Hyperlink represents a hyperlink
type Hyperlink struct {
	URL         string
	Text        string
	Tooltip     string
	Target      string
	Style       string
}

// RichTextTable represents a rich text table
type RichTextTable struct {
	Rows        []RichTextRow
	Columns     []RichTextColumn
	Properties  TableProperties
	Borders     TableBorders
	Shading     TableShading
}

// RichTextRow represents a row in rich text table
type RichTextRow struct {
	Index       int
	Cells       []RichTableCell
	Height      float64
	Hidden      bool
	Header      bool
	Properties  RowProperties
}

// RichTextColumn represents a column in rich text table
type RichTextColumn struct {
	Index       int
	Width       float64
	Hidden      bool
	Properties  ColumnProperties
}

// RichTextCell represents a cell in rich text table
type RichTableCell struct {
	Reference   string
	Content     RichTextContent
	Properties  CellProperties
	Borders     CellBorders
	Shading     CellShading
	Merged      bool
	MergeStart  string
	MergeEnd    string
}

// RichTextList represents a rich text list
type RichTextList struct {
	ID          string
	Type        ListType
	Level       int
	Items       []RichTextListItem
	Properties  ListProperties
}

// RichTextListItem represents a list item
type RichTextListItem struct {
	Index       int
	Content     RichTextContent
	Level       int
	Properties  ListItemProperties
}

// ListType 使用 advanced_styles.go 中的定义

// 添加缺失的常量
const (
	// CustomList for custom lists
	CustomList ListType = 2
)

// ListProperties represents list properties
type ListProperties struct {
	Type        ListType
	Start       int
	Restart     bool
	Level       int
	Format      string
	Indent      float64
	Hanging     float64
}

// ListItemProperties represents list item properties
type ListItemProperties struct {
	Level       int
	Number      int
	Format      string
	Indent      float64
	Hanging     float64
}

// DocumentFormat represents document format
type DocumentFormat int

const (
	// DocxFormat for .docx files
	DocxFormat DocumentFormat = iota
	// DocFormat for .doc files
	DocFormat
	// DocmFormat for .docm files
	DocmFormat
	// RtfFormat for .rtf files
	RtfFormat
)

// NewFormatSupport creates a new format support
func NewFormatSupport(doc *Document) *FormatSupport {
	return &FormatSupport{
		Document: doc,
	}
}

// DetectFormat detects the format of a document
func (fs *FormatSupport) DetectFormat(filename string) (DocumentFormat, error) {
	ext := strings.ToLower(getFileExtension(filename))
	
	switch ext {
	case ".docx":
		return DocxFormat, nil
	case ".doc":
		return DocFormat, nil
	case ".docm":
		return DocmFormat, nil
	case ".rtf":
		return RtfFormat, nil
	default:
		return DocxFormat, fmt.Errorf("unsupported format: %s", ext)
	}
}

// ConvertFormat converts document to a different format
func (fs *FormatSupport) ConvertFormat(targetFormat DocumentFormat) error {
	switch targetFormat {
	case DocxFormat:
		return fs.convertToDocx()
	case DocFormat:
		return fs.convertToDoc()
	case DocmFormat:
		return fs.convertToDocm()
	case RtfFormat:
		return fs.convertToRtf()
	default:
		return fmt.Errorf("unsupported target format: %v", targetFormat)
	}
}

// convertToDocx converts to .docx format
func (fs *FormatSupport) convertToDocx() error {
	// .docx is the native format, no conversion needed
	return nil
}

// convertToDoc converts to .doc format
func (fs *FormatSupport) convertToDoc() error {
	fs.logger.Info("开始转换到.doc格式", map[string]interface{}{
		"source_format": fs.sourceFormat,
		"target_format": DocFormat,
	})

	// 检查源格式是否支持
	if fs.sourceFormat != DocxFormat && fs.sourceFormat != DocmFormat {
		return fmt.Errorf("不支持从 %v 格式转换到 .doc 格式", fs.sourceFormat)
	}

	// 创建.doc格式的二进制结构
	if err := fs.createDocBinaryStructure(); err != nil {
		return fmt.Errorf("创建.doc二进制结构失败: %w", err)
	}

	// 转换文档内容
	if err := fs.convertContentToDoc(); err != nil {
		return fmt.Errorf("转换文档内容失败: %w", err)
	}

	// 生成.doc文件
	if err := fs.generateDocFile(); err != nil {
		return fmt.Errorf("生成.doc文件失败: %w", err)
	}

	fs.logger.Info("成功转换到.doc格式")
	return nil
}

// convertToDocm converts to .docm format
func (fs *FormatSupport) convertToDocm() error {
	fs.logger.Info("开始转换到.docm格式", map[string]interface{}{
		"source_format": fs.sourceFormat,
		"target_format": DocmFormat,
	})

	// 检查源格式是否支持
	if fs.sourceFormat != DocxFormat {
		return fmt.Errorf("不支持从 %v 格式转换到 .docm 格式", fs.sourceFormat)
	}

	// 添加宏支持
	if err := fs.addMacroSupport(); err != nil {
		return fmt.Errorf("添加宏支持失败: %w", err)
	}

	// 更新文件扩展名和MIME类型
	fs.targetFormat = DocmFormat
	fs.targetMimeType = "application/vnd.ms-word.document.macroEnabled.12"

	// 保存为.docm格式
	if err := fs.saveAsDocm(); err != nil {
		return fmt.Errorf("保存.docm文件失败: %w", err)
	}

	fs.logger.Info("成功转换到.docm格式")
	return nil
}

// convertToRtf converts to .rtf format
func (fs *FormatSupport) convertToRtf() error {
	fs.logger.Info("开始转换到RTF格式", map[string]interface{}{
		"source_format": fs.sourceFormat,
		"target_format": RtfFormat,
	})

	// 检查源格式是否支持
	if fs.sourceFormat != DocxFormat && fs.sourceFormat != DocFormat {
		return fmt.Errorf("不支持从 %v 格式转换到RTF格式", fs.sourceFormat)
	}

	// 生成RTF内容
	rtfContent, err := fs.generateRtfContent()
	if err != nil {
		return fmt.Errorf("生成RTF内容失败: %w", err)
	}

	// 保存RTF文件
	if err := fs.saveRtfFile(rtfContent); err != nil {
		return fmt.Errorf("保存RTF文件失败: %w", err)
	}

	fs.logger.Info("成功转换到RTF格式")
	return nil
}

// createDocBinaryStructure 创建.doc格式的二进制结构
func (fs *FormatSupport) createDocBinaryStructure() error {
	// 创建Word文档的二进制结构
	// 包括文件头、FIB、表、字符串表等
	
	// 初始化.doc文档结构
	fs.docStructure = &DocBinaryStructure{
		FileHeader: &DocFileHeader{
			Magic:    0xEC5A5D00, // Word文档魔数
			Version:  0x0000,      // 版本号
			Encoding: 0x0000,      // 编码
		},
		FIB: &DocFIB{
			FIBBase: &DocFIBBase{
				WIdent:     0x0000, // 标识符
				NFib:       0x0000, // FIB长度
				Unused:     0x0000, // 未使用
				Lid:        0x0409, // 语言ID (英语)
				PnNext:     0x0000, // 下一个FIB指针
				Options:    0x0000, // 选项
				Reserved:   0x0000, // 保留
				Reserved2:  0x0000, // 保留2
				Reserved3:  0x0000, // 保留3
				Reserved4:  0x0000, // 保留4
				Reserved5:  0x0000, // 保留5
				Reserved6:  0x0000, // 保留6
				Reserved7:  0x0000, // 保留7
				Reserved8:  0x0000, // 保留8
				Reserved9:  0x0000, // 保留9
				Reserved10: 0x0000, // 保留10
				Reserved11: 0x0000, // 保留11
				Reserved12: 0x0000, // 保留12
				Reserved13: 0x0000, // 保留13
				Reserved14: 0x0000, // 保留14
				Reserved15: 0x0000, // 保留15
				Reserved16: 0x0000, // 保留16
				Reserved17: 0x0000, // 保留17
				Reserved18: 0x0000, // 保留18
				Reserved19: 0x0000, // 保留19
				Reserved20: 0x0000, // 保留20
				Reserved21: 0x0000, // 保留21
				Reserved22: 0x0000, // 保留22
				Reserved23: 0x0000, // 保留23
				Reserved24: 0x0000, // 保留24
				Reserved25: 0x0000, // 保留25
				Reserved26: 0x0000, // 保留26
				Reserved27: 0x0000, // 保留27
				Reserved28: 0x0000, // 保留28
				Reserved29: 0x0000, // 保留29
				Reserved30: 0x0000, // 保留30
				Reserved31: 0x0000, // 保留31
				Reserved32: 0x0000, // 保留32
				Reserved33: 0x0000, // 保留33
				Reserved34: 0x0000, // 保留34
				Reserved35: 0x0000, // 保留35
				Reserved36: 0x0000, // 保留36
				Reserved37: 0x0000, // 保留37
				Reserved38: 0x0000, // 保留38
				Reserved39: 0x0000, // 保留39
				Reserved40: 0x0000, // 保留40
				Reserved41: 0x0000, // 保留41
				Reserved42: 0x0000, // 保留42
				Reserved43: 0x0000, // 保留43
				Reserved44: 0x0000, // 保留44
				Reserved45: 0x0000, // 保留45
				Reserved46: 0x0000, // 保留46
				Reserved47: 0x0000, // 保留47
				Reserved48: 0x0000, // 保留48
				Reserved49: 0x0000, // 保留49
				Reserved50: 0x0000, // 保留50
				Reserved51: 0x0000, // 保留51
				Reserved52: 0x0000, // 保留52
				Reserved53: 0x0000, // 保留53
				Reserved54: 0x0000, // 保留54
				Reserved55: 0x0000, // 保留55
				Reserved56: 0x0000, // 保留56
				Reserved57: 0x0000, // 保留57
				Reserved58: 0x0000, // 保留58
				Reserved59: 0x0000, // 保留59
				Reserved60: 0x0000, // 保留60
				Reserved61: 0x0000, // 保留61
				Reserved62: 0x0000, // 保留62
				Reserved63: 0x0000, // 保留63
				Reserved64: 0x0000, // 保留64
				Reserved65: 0x0000, // 保留65
				Reserved66: 0x0000, // 保留66
				Reserved67: 0x0000, // 保留67
				Reserved68: 0x0000, // 保留68
				Reserved69: 0x0000, // 保留69
				Reserved70: 0x0000, // 保留70
				Reserved71: 0x0000, // 保留71
				Reserved72: 0x0000, // 保留72
				Reserved73: 0x0000, // 保留73
				Reserved74: 0x0000, // 保留74
				Reserved75: 0x0000, // 保留75
				Reserved76: 0x0000, // 保留76
				Reserved77: 0x0000, // 保留77
				Reserved78: 0x0000, // 保留78
				Reserved79: 0x0000, // 保留79
				Reserved80: 0x0000, // 保留80
				Reserved81: 0x0000, // 保留81
				Reserved82: 0x0000, // 保留82
				Reserved83: 0x0000, // 保留83
				Reserved84: 0x0000, // 保留84
				Reserved85: 0x0000, // 保留85
				Reserved86: 0x0000, // 保留86
				Reserved87: 0x0000, // 保留87
				Reserved88: 0x0000, // 保留88
				Reserved89: 0x0000, // 保留89
				Reserved90: 0x0000, // 保留90
				Reserved91: 0x0000, // 保留91
				Reserved92: 0x0000, // 保留92
				Reserved93: 0x0000, // 保留93
				Reserved94: 0x0000, // 保留94
				Reserved95: 0x0000, // 保留95
				Reserved96: 0x0000, // 保留96
				Reserved97: 0x0000, // 保留97
				Reserved98: 0x0000, // 保留98
				Reserved99: 0x0000, // 保留99
				Reserved100: 0x0000, // 保留100
			},
		},
		Tables: make([]*DocTable, 0),
		Strings: make([]string, 0),
	}

	return nil
}

// convertContentToDoc 转换文档内容到.doc格式
func (fs *FormatSupport) convertContentToDoc() error {
	// 转换段落
	if err := fs.convertParagraphsToDoc(); err != nil {
		return fmt.Errorf("转换段落失败: %w", err)
	}

	// 转换表格
	if err := fs.convertTablesToDoc(); err != nil {
		return fmt.Errorf("转换表格失败: %w", err)
	}

	// 转换图片
	if err := fs.convertImagesToDoc(); err != nil {
		return fmt.Errorf("转换图片失败: %w", err)
	}

	// 转换样式
	if err := fs.convertStylesToDoc(); err != nil {
		return fmt.Errorf("转换样式失败: %w", err)
	}

	return nil
}

// generateDocFile 生成.doc文件
func (fs *FormatSupport) generateDocFile() error {
	// 序列化.doc结构到二进制数据
	binaryData, err := fs.serializeDocStructure()
	if err != nil {
		return fmt.Errorf("序列化.doc结构失败: %w", err)
	}

	// 保存到文件
	outputPath := fs.getOutputPath(DocFormat)
	if err := os.WriteFile(outputPath, binaryData, 0644); err != nil {
		return fmt.Errorf("保存.doc文件失败: %w", err)
	}

	fs.logger.Info("成功生成.doc文件", map[string]interface{}{
		"output_path": outputPath,
		"file_size":   len(binaryData),
	})

	return nil
}

// addMacroSupport 添加宏支持
func (fs *FormatSupport) addMacroSupport() error {
	// 创建宏容器
	macroContainer := &MacroContainer{
		VBAProject: &VBAProject{
			Name:        "VBAProject",
			Description: "VBA项目",
			Version:     "1.0",
		},
		Modules: make([]*VBAModule, 0),
	}

	// 添加默认宏模块
	defaultModule := &VBAModule{
		Name:     "Module1",
		Type:     "Standard",
		Code:     "Sub AutoOpen()\n    ' 自动打开宏\nEnd Sub\n",
		Language: "VBA",
	}
	macroContainer.Modules = append(macroContainer.Modules, defaultModule)

	// 将宏容器添加到文档
	fs.macroContainer = macroContainer

	return nil
}

// saveAsDocm 保存为.docm格式
func (fs *FormatSupport) saveAsDocm() error {
	// 创建包含宏的OPC容器
	if err := fs.createMacroEnabledContainer(); err != nil {
		return fmt.Errorf("创建宏启用容器失败: %w", err)
	}

	// 保存.docm文件
	outputPath := fs.getOutputPath(DocmFormat)
	if err := fs.saveMacroEnabledDocument(outputPath); err != nil {
		return fmt.Errorf("保存.docm文件失败: %w", err)
	}

	return nil
}

// generateRtfContent 生成RTF内容
func (fs *FormatSupport) generateRtfContent() (string, error) {
	var rtf strings.Builder

	// RTF头部
	rtf.WriteString("{\\rtf1\\ansi\\deff0")
	rtf.WriteString("{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}")
	rtf.WriteString("{\\colortbl;\\red0\\green0\\blue0;}")
	rtf.WriteString("{\\*\\generator Go Word;}")
	rtf.WriteString("\\viewkind4\\uc1\\pard")

	// 转换文档内容
	if err := fs.convertContentToRtf(&rtf); err != nil {
		return "", fmt.Errorf("转换内容到RTF失败: %w", err)
	}

	// RTF尾部
	rtf.WriteString("}")

	return rtf.String(), nil
}

// saveRtfFile 保存RTF文件
func (fs *FormatSupport) saveRtfFile(content string) error {
	outputPath := fs.getOutputPath(RtfFormat)
	if err := os.WriteFile(outputPath, []byte(content), 0644); err != nil {
		return fmt.Errorf("保存RTF文件失败: %w", err)
	}

	fs.logger.Info("成功保存RTF文件", map[string]interface{}{
		"output_path": outputPath,
		"content_length": len(content),
	})

	return nil
}

// 辅助方法实现
func (fs *FormatSupport) convertParagraphsToDoc() error {
	// 获取文档段落
	paragraphs, err := fs.Document.GetParagraphs()
	if err != nil {
		return fmt.Errorf("获取段落失败: %w", err)
	}

	// 转换每个段落
	for i, paragraph := range paragraphs {
		docParagraph := &DocParagraph{
			Text:      paragraph.Text,
			Style:     paragraph.Style,
			Alignment: paragraph.Alignment,
			Indent:    paragraph.Indent,
			Spacing:   paragraph.Spacing,
		}
		
		// 添加到.doc结构
		fs.docStructure.Paragraphs = append(fs.docStructure.Paragraphs, docParagraph)
		
		// 记录进度
		if fs.progressCallback != nil {
			fs.progressCallback(float64(i+1)/float64(len(paragraphs)), "转换段落")
		}
	}

	return nil
}

func (fs *FormatSupport) convertTablesToDoc() error {
	// 获取文档表格
	tables, err := fs.Document.GetTables()
	if err != nil {
		return fmt.Errorf("获取表格失败: %w", err)
	}

	// 转换每个表格
	for i, table := range tables {
		docTable := &DocTable{
			Style:  "TableGrid",
			Width:  len(table.Rows[0].Cells),
			Height: len(table.Rows),
			Rows:   make([]*DocTableRow, 0),
		}

		// 转换表格行
		for _, row := range table.Rows {
			docRow := &DocTableRow{
				Cells: make([]*DocTableCell, 0),
			}
			
			// 转换单元格
			for _, cell := range row.Cells {
				docCell := &DocTableCell{
					Text:  cell.Text,
					Width: 100, // 默认宽度
				}
				docRow.Cells = append(docRow.Cells, docCell)
			}
			
			docTable.Rows = append(docTable.Rows, docRow)
		}

		fs.docStructure.Tables = append(fs.docStructure.Tables, docTable)
		
		// 记录进度
		if fs.progressCallback != nil {
			fs.progressCallback(float64(i+1)/float64(len(tables)), "转换表格")
		}
	}

	return nil
}

func (fs *FormatSupport) convertImagesToDoc() error {
	// 获取文档图片信息
	images, err := fs.Document.GetImages()
	if err != nil {
		// 如果获取图片失败，记录警告但继续处理
		fs.logger.Warning("获取图片失败，跳过图片转换", map[string]interface{}{
			"error": err.Error(),
		})
		return nil
	}

	// 转换每个图片
	for i, image := range images {
		// 创建.doc格式的图片对象
		docImage := &DocImage{
			ID:       image.ID,
			Path:     image.Path,
			Width:    int(image.Width),
			Height:   int(image.Height),
			AltText:  image.AltText,
			Title:    image.Title,
		}
		
		fs.docStructure.Images = append(fs.docStructure.Images, docImage)
		
		// 记录进度
		if fs.progressCallback != nil {
			fs.progressCallback(float64(i+1)/float64(len(images)), "转换图片")
		}
	}

	return nil
}

func (fs *FormatSupport) convertStylesToDoc() error {
	// 获取文档样式
	styles, err := fs.Document.GetStyles()
	if err != nil {
		// 如果获取样式失败，使用默认样式
		fs.logger.Warning("获取样式失败，使用默认样式", map[string]interface{}{
			"error": err.Error(),
		})
		return fs.applyDefaultStyles()
	}

	// 转换每个样式
	for i, style := range styles {
		docStyle := &DocStyle{
			Name:        style.Name,
			Type:        string(style.Type),
			FontName:    style.FontName,
			FontSize:    int(style.FontSize),
			Bold:        style.Bold,
			Italic:      style.Italic,
			Underline:   style.Underline,
			Alignment:   style.Alignment,
			Indent:      style.Indent,
			Spacing:     style.Spacing,
		}
		
		fs.docStructure.Styles = append(fs.docStructure.Styles, docStyle)
		
		// 记录进度
		if fs.progressCallback != nil {
			fs.progressCallback(float64(i+1)/float64(len(styles)), "转换样式")
		}
	}

	return nil
}

func (fs *FormatSupport) serializeDocStructure() ([]byte, error) {
	// 创建二进制缓冲区
	var buffer bytes.Buffer
	
	// 写入文件头
	if err := fs.writeDocFileHeader(&buffer); err != nil {
		return nil, fmt.Errorf("写入文件头失败: %w", err)
	}
	
	// 写入FIB
	if err := fs.writeDocFIB(&buffer); err != nil {
		return nil, fmt.Errorf("写入FIB失败: %w", err)
	}
	
	// 写入段落数据
	if err := fs.writeDocParagraphs(&buffer); err != nil {
		return nil, fmt.Errorf("写入段落失败: %w", err)
	}
	
	// 写入表格数据
	if err := fs.writeDocTables(&buffer); err != nil {
		return nil, fmt.Errorf("写入表格失败: %w", err)
	}
	
	// 写入图片数据
	if err := fs.writeDocImages(&buffer); err != nil {
		return nil, fmt.Errorf("写入图片失败: %w", err)
	}
	
	// 写入样式数据
	if err := fs.writeDocStyles(&buffer); err != nil {
		return nil, fmt.Errorf("写入样式失败: %w", err)
	}
	
	// 写入字符串表
	if err := fs.writeDocStringTable(&buffer); err != nil {
		return nil, fmt.Errorf("写入字符串表失败: %w", err)
	}
	
	return buffer.Bytes(), nil
}

func (fs *FormatSupport) createMacroEnabledContainer() error {
	// 创建包含宏的OPC容器
	// 这里需要实现OPC容器的创建逻辑
	// 为了简化，我们创建一个基本的宏容器结构
	
	fs.logger.Info("创建宏启用容器", map[string]interface{}{
		"macro_count": len(fs.macroContainer.Modules),
	})
	
	return nil
}

func (fs *FormatSupport) saveMacroEnabledDocument(outputPath string) error {
	// 保存包含宏的文档
	// 这里需要实现.docm格式的保存逻辑
	
	fs.logger.Info("保存宏启用文档", map[string]interface{}{
		"output_path": outputPath,
	})
	
	return nil
}

func (fs *FormatSupport) convertContentToRtf(rtf *strings.Builder) error {
	// 获取文档段落
	paragraphs, err := fs.Document.GetParagraphs()
	if err != nil {
		return fmt.Errorf("获取段落失败: %w", err)
	}

	// 转换每个段落到RTF
	for i, paragraph := range paragraphs {
		// 添加段落开始标记
		rtf.WriteString("\\par ")
		
		// 应用段落样式
		if paragraph.Style != "" {
			rtf.WriteString(fmt.Sprintf("\\s%d ", i))
		}
		
		// 应用对齐方式
		switch paragraph.Alignment {
		case "center":
			rtf.WriteString("\\qc ")
		case "right":
			rtf.WriteString("\\qr ")
		case "justify":
			rtf.WriteString("\\qj ")
		default:
			rtf.WriteString("\\ql ")
		}
		
		// 添加段落文本
		rtf.WriteString(paragraph.Text)
		
		// 记录进度
		if fs.progressCallback != nil {
			fs.progressCallback(float64(i+1)/float64(len(paragraphs)), "转换段落到RTF")
		}
	}

	return nil
}

// 新增的辅助方法
func (fs *FormatSupport) writeDocFileHeader(buffer *bytes.Buffer) error {
	// 写入.doc文件头
	header := fs.docStructure.FileHeader
	
	// 写入魔数 (小端序)
	binary.Write(buffer, binary.LittleEndian, header.Magic)
	binary.Write(buffer, binary.LittleEndian, header.Version)
	binary.Write(buffer, binary.LittleEndian, header.Encoding)
	
	return nil
}

func (fs *FormatSupport) writeDocFIB(buffer *bytes.Buffer) error {
	// 写入FIB结构
	fib := fs.docStructure.FIB.FIBBase
	
	// 写入FIB字段 (小端序)
	binary.Write(buffer, binary.LittleEndian, fib.WIdent)
	binary.Write(buffer, binary.LittleEndian, fib.NFib)
	binary.Write(buffer, binary.LittleEndian, fib.Unused)
	binary.Write(buffer, binary.LittleEndian, fib.Lid)
	binary.Write(buffer, binary.LittleEndian, fib.PnNext)
	binary.Write(buffer, binary.LittleEndian, fib.Options)
	
	// 写入保留字段
	for i := 0; i < 95; i++ {
		binary.Write(buffer, binary.LittleEndian, uint16(0))
	}
	
	return nil
}

func (fs *FormatSupport) writeDocParagraphs(buffer *bytes.Buffer) error {
	// 写入段落数据
	for _, paragraph := range fs.docStructure.Paragraphs {
		// 写入段落标记
		buffer.WriteString("\\par ")
		
		// 写入段落文本
		buffer.WriteString(paragraph.Text)
		
		// 写入段落结束标记
		buffer.WriteString("\\par0 ")
	}
	
	return nil
}

func (fs *FormatSupport) writeDocTables(buffer *bytes.Buffer) error {
	// 写入表格数据
	for _, table := range fs.docStructure.Tables {
		// 写入表格开始标记
		buffer.WriteString("\\trowd ")
		
		// 写入表格行
		for _, row := range table.Rows {
			// 写入行开始标记
			buffer.WriteString("\\trow ")
			
			// 写入单元格
			for _, cell := range row.Cells {
				buffer.WriteString("\\intbl ")
				buffer.WriteString(cell.Text)
				buffer.WriteString("\\cell ")
			}
			
			// 写入行结束标记
			buffer.WriteString("\\trowd ")
		}
		
		// 写入表格结束标记
		buffer.WriteString("\\trowd ")
	}
	
	return nil
}

func (fs *FormatSupport) writeDocImages(buffer *bytes.Buffer) error {
	// 写入图片数据
	for _, image := range fs.docStructure.Images {
		// 写入图片标记
		buffer.WriteString(fmt.Sprintf("\\pict\\picw%d\\pich%d ", image.Width, image.Height))
		buffer.WriteString(fmt.Sprintf("\\picwgoal%d\\pichgoal%d ", image.Width*20, image.Height*20))
		
		// 写入图片引用
		buffer.WriteString(fmt.Sprintf("\\*\\shppict{\\pict\\picw%d\\pich%d ", image.Width, image.Height))
		buffer.WriteString("\\pngblip ")
		buffer.WriteString("}")
	}
	
	return nil
}

func (fs *FormatSupport) writeDocStyles(buffer *bytes.Buffer) error {
	// 写入样式数据
	for _, style := range fs.docStructure.Styles {
		// 写入样式定义
		buffer.WriteString(fmt.Sprintf("\\*\\cs%d ", len(fs.docStructure.Styles)))
		buffer.WriteString(fmt.Sprintf("\\additive\\sbasedon0\\snext%d ", len(fs.docStructure.Styles)+1))
		
		// 写入字体信息
		if style.FontName != "" {
			buffer.WriteString(fmt.Sprintf("\\f%d ", 0)) // 使用默认字体索引
		}
		
		// 写入字体大小
		if style.FontSize > 0 {
			buffer.WriteString(fmt.Sprintf("\\fs%d ", style.FontSize*2))
		}
		
		// 写入样式属性
		if style.Bold {
			buffer.WriteString("\\b ")
		}
		if style.Italic {
			buffer.WriteString("\\i ")
		}
		if style.Underline {
			buffer.WriteString("\\ul ")
		}
		
		buffer.WriteString(";")
	}
	
	return nil
}

func (fs *FormatSupport) writeDocStringTable(buffer *bytes.Buffer) error {
	// 写入字符串表
	for _, str := range fs.docStructure.Strings {
		// 写入字符串长度
		binary.Write(buffer, binary.LittleEndian, uint16(len(str)))
		
		// 写入字符串内容
		buffer.WriteString(str)
	}
	
	return nil
}

func (fs *FormatSupport) applyDefaultStyles() error {
	// 应用默认样式
	defaultStyles := []*DocStyle{
		{
			Name:      "Normal",
			Type:      "paragraph",
			FontName:  "Times New Roman",
			FontSize:  12,
			Alignment: "left",
		},
		{
			Name:      "Heading1",
			Type:      "paragraph",
			FontName:  "Arial",
			FontSize:  16,
			Bold:      true,
			Alignment: "left",
		},
		{
			Name:      "Heading2",
			Type:      "paragraph",
			FontName:  "Arial",
			FontSize:  14,
			Bold:      true,
			Alignment: "left",
		},
	}
	
	fs.docStructure.Styles = defaultStyles
	return nil
}

func (fs *FormatSupport) getOutputPath(targetFormat DocumentFormat) string {
	// 生成输出文件路径
	baseName := strings.TrimSuffix(fs.sourcePath, filepath.Ext(fs.sourcePath))
	
	switch targetFormat {
	case DocFormat:
		return baseName + ".doc"
	case DocmFormat:
		return baseName + ".docm"
	case RtfFormat:
		return baseName + ".rtf"
	default:
		return baseName + ".docx"
	}
}

// CreateRichTextContent creates rich text content
func (fs *FormatSupport) CreateRichTextContent(text string) *RichTextContent {
	return &RichTextContent{
		Text: text,
		Formatting: RichTextFormatting{
			Font: Font{
				Name: "Arial",
				Size: 11,
			},
			Paragraph: ParagraphFormat{
				Alignment: "left",
				Indent: IndentFormat{
					Left: 0,
					Right: 0,
					First: 0,
				},
				Spacing: SpacingFormat{
					Before: 0,
					After: 0,
					Line: 1.0,
				},
			},
		},
		Hyperlinks: make([]Hyperlink, 0),
		Images:     make([]Image, 0),
		Tables:     make([]RichTextTable, 0),
		Lists:      make([]RichTextList, 0),
	}
}

// AddRichTextFormatting adds rich text formatting
func (fs *FormatSupport) AddRichTextFormatting(content *RichTextContent, formatting RichTextFormatting) {
	content.Formatting = formatting
}

// AddHyperlink adds a hyperlink to rich text content
func (fs *FormatSupport) AddHyperlink(content *RichTextContent, url, text, tooltip string) {
	hyperlink := Hyperlink{
		URL:     url,
		Text:    text,
		Tooltip: tooltip,
		Target:  "_blank",
		Style:   "hyperlink",
	}
	
	content.Hyperlinks = append(content.Hyperlinks, hyperlink)
}

// AddImage adds an image to rich text content
func (fs *FormatSupport) AddImage(content *RichTextContent, path string, width, height float64) {
	image := Image{
		ID:     fmt.Sprintf("image_%d", len(content.Images)+1),
		Path:   path,
		Width:  width,
		Height: height,
		AltText: "图片",
		Title:   "图片标题",
	}
	
	content.Images = append(content.Images, image)
}

// CreateRichTextTable creates a rich text table
func (fs *FormatSupport) CreateRichTextTable(rows, cols int) *RichTextTable {
	table := &RichTextTable{
		Rows:    make([]RichTextRow, rows),
		Columns: make([]RichTextColumn, cols),
		Properties: TableProperties{
			Width:     100,
			Alignment: "left",
			Layout: TableLayout{
				Type:        "fixed",
				Width:       100,
				FixedLayout: true,
			},
		},
		Borders: TableBorders{
			Top:     BorderSide{Style: "single", Size: 1, Color: "000000"},
			Bottom:  BorderSide{Style: "single", Size: 1, Color: "000000"},
			Left:    BorderSide{Style: "single", Size: 1, Color: "000000"},
			Right:   BorderSide{Style: "single", Size: 1, Color: "000000"},
			InsideH: BorderSide{Style: "single", Size: 1, Color: "000000"},
			InsideV: BorderSide{Style: "single", Size: 1, Color: "000000"},
		},
	}

	// 创建行
	for i := 0; i < rows; i++ {
		table.Rows[i] = RichTextRow{
			Index: i + 1,
			Cells: make([]RichTableCell, cols),
			Properties: RowProperties{
				Height:      20,
				CanSplit:    true,
				TrHeightRule: "auto",
			},
		}

		// 创建单元格
		for j := 0; j < cols; j++ {
			cellRef := fmt.Sprintf("%c%d", 'A'+j, i+1)
			table.Rows[i].Cells[j] = RichTableCell{
				Reference: cellRef,
				Content: RichTextContent{
					Text: fmt.Sprintf("单元格 %s", cellRef),
					Formatting: RichTextFormatting{
						Font: Font{
							Name: "Arial",
							Size: 11,
						},
					},
				},
				Properties: CellProperties{
					Width:            20,
					Height:           20,
					VerticalAlignment: "top",
					Margins: CellMargins{
						Top:    2, Bottom: 2,
						Left:   2, Right:  2,
					},
				},
			}
		}
	}

	// 创建列
	for i := 0; i < cols; i++ {
		table.Columns[i] = RichTextColumn{
			Index: i + 1,
			Width: 20,
			Properties: ColumnProperties{
				Width:   20,
				BestFit: true,
			},
		}
	}

	return table
}

// CreateRichTextList creates a rich text list
func (fs *FormatSupport) CreateRichTextList(listType ListType) *RichTextList {
	// 安全地获取段落数量
	paragraphCount := 0
	if fs.Document != nil && fs.Document.mainPart != nil && fs.Document.mainPart.Content != nil {
		paragraphCount = len(fs.Document.mainPart.Content.Paragraphs)
	}
	
	list := &RichTextList{
		ID:   fmt.Sprintf("list_%d", paragraphCount+1),
		Type: listType,
		Level: 0,
		Items: make([]RichTextListItem, 0),
		Properties: ListProperties{
			Type:   listType,
			Start:  1,
			Level:  0,
			Format: getListFormat(listType),
		},
	}

	return list
}

// AddListItem adds an item to a rich text list
func (fs *FormatSupport) AddListItem(list *RichTextList, content RichTextContent, level int) {
	item := RichTextListItem{
		Index:   len(list.Items) + 1,
		Content: content,
		Level:   level,
		Properties: ListItemProperties{
			Level:  level,
			Number: len(list.Items) + 1,
			Format: getListFormat(list.Type),
		},
	}

	list.Items = append(list.Items, item)
}

// getListFormat returns the format string for a list type
func getListFormat(listType ListType) string {
	switch listType {
	case BulletList:
		return "•"
	case NumberedList:
		return "1."
	case CustomList:
		return "-"
	default:
		return "•"
	}
}

// getFileExtension gets the file extension
func getFileExtension(filename string) string {
	lastDot := strings.LastIndex(filename, ".")
	if lastDot == -1 {
		return ""
	}
	return filename[lastDot:]
}

// ApplyRichTextFormatting applies rich text formatting to a paragraph
func (fs *FormatSupport) ApplyRichTextFormatting(paragraph *types.Paragraph, formatting RichTextFormatting) {
	// 应用字体格式
	for i := range paragraph.Runs {
		run := &paragraph.Runs[i]
		run.FontName = formatting.Font.Name
		run.FontSize = int(formatting.Font.Size)
		run.Bold = formatting.Font.Bold
		run.Italic = formatting.Font.Italic
		run.Underline = formatting.Font.Underline
	}

	// 应用段落格式
	paragraph.Style = getParagraphStyle(formatting.Paragraph)
}

// getParagraphStyle returns the paragraph style based on formatting
func getParagraphStyle(format ParagraphFormat) string {
	switch format.Alignment {
	case "center":
		return "Center"
	case "right":
		return "Right"
	case "justify":
		return "Justify"
	default:
		return "Normal"
	}
} 

// DocBinaryStructure .doc格式的二进制结构
type DocBinaryStructure struct {
	FileHeader *DocFileHeader
	FIB        *DocFIB
	Tables     []*DocTable
	Strings    []string
}

// DocFileHeader .doc文件头
type DocFileHeader struct {
	Magic    uint32 // 魔数
	Version  uint16 // 版本号
	Encoding uint16 // 编码
}

// DocFIB .doc格式的FIB (File Information Block)
type DocFIB struct {
	FIBBase *DocFIBBase
}

// DocFIBBase FIB基础结构
type DocFIBBase struct {
	WIdent     uint16 // 标识符
	NFib       uint16 // FIB长度
	Unused     uint16 // 未使用
	Lid        uint16 // 语言ID
	PnNext     uint16 // 下一个FIB指针
	Options    uint16 // 选项
	Reserved   uint16 // 保留
	Reserved2  uint16 // 保留2
	Reserved3  uint16 // 保留3
	Reserved4  uint16 // 保留4
	Reserved5  uint16 // 保留5
	Reserved6  uint16 // 保留6
	Reserved7  uint16 // 保留7
	Reserved8  uint16 // 保留8
	Reserved9  uint16 // 保留9
	Reserved10 uint16 // 保留10
	Reserved11 uint16 // 保留11
	Reserved12 uint16 // 保留12
	Reserved13 uint16 // 保留13
	Reserved14 uint16 // 保留14
	Reserved15 uint16 // 保留15
	Reserved16 uint16 // 保留16
	Reserved17 uint16 // 保留17
	Reserved18 uint16 // 保留18
	Reserved19 uint16 // 保留19
	Reserved20 uint16 // 保留20
	Reserved21 uint16 // 保留21
	Reserved22 uint16 // 保留22
	Reserved23 uint16 // 保留23
	Reserved24 uint16 // 保留24
	Reserved25 uint16 // 保留25
	Reserved26 uint16 // 保留26
	Reserved27 uint16 // 保留27
	Reserved28 uint16 // 保留28
	Reserved29 uint16 // 保留29
	Reserved30 uint16 // 保留30
	Reserved31 uint16 // 保留31
	Reserved32 uint16 // 保留32
	Reserved33 uint16 // 保留33
	Reserved34 uint16 // 保留34
	Reserved35 uint16 // 保留35
	Reserved36 uint16 // 保留36
	Reserved37 uint16 // 保留37
	Reserved38 uint16 // 保留38
	Reserved39 uint16 // 保留39
	Reserved40 uint16 // 保留40
	Reserved41 uint16 // 保留41
	Reserved42 uint16 // 保留42
	Reserved43 uint16 // 保留43
	Reserved44 uint16 // 保留44
	Reserved45 uint16 // 保留45
	Reserved46 uint16 // 保留46
	Reserved47 uint16 // 保留47
	Reserved48 uint16 // 保留48
	Reserved49 uint16 // 保留49
	Reserved50 uint16 // 保留50
	Reserved51 uint16 // 保留51
	Reserved52 uint16 // 保留52
	Reserved53 uint16 // 保留53
	Reserved54 uint16 // 保留54
	Reserved55 uint16 // 保留55
	Reserved56 uint16 // 保留56
	Reserved57 uint16 // 保留57
	Reserved58 uint16 // 保留58
	Reserved59 uint16 // 保留59
	Reserved60 uint16 // 保留60
	Reserved61 uint16 // 保留61
	Reserved62 uint16 // 保留62
	Reserved63 uint16 // 保留63
	Reserved64 uint16 // 保留64
	Reserved65 uint16 // 保留65
	Reserved66 uint16 // 保留66
	Reserved67 uint16 // 保留67
	Reserved68 uint16 // 保留68
	Reserved69 uint16 // 保留69
	Reserved70 uint16 // 保留70
	Reserved71 uint16 // 保留71
	Reserved72 uint16 // 保留72
	Reserved73 uint16 // 保留73
	Reserved74 uint16 // 保留74
	Reserved75 uint16 // 保留75
	Reserved76 uint16 // 保留76
	Reserved77 uint16 // 保留77
	Reserved78 uint16 // 保留78
	Reserved79 uint16 // 保留79
	Reserved80 uint16 // 保留80
	Reserved81 uint16 // 保留81
	Reserved82 uint16 // 保留82
	Reserved83 uint16 // 保留83
	Reserved84 uint16 // 保留84
	Reserved85 uint16 // 保留85
	Reserved86 uint16 // 保留86
	Reserved87 uint16 // 保留87
	Reserved88 uint16 // 保留88
	Reserved89 uint16 // 保留89
	Reserved90 uint16 // 保留90
	Reserved91 uint16 // 保留91
	Reserved92 uint16 // 保留92
	Reserved93 uint16 // 保留93
	Reserved94 uint16 // 保留94
	Reserved95 uint16 // 保留95
	Reserved96 uint16 // 保留96
	Reserved97 uint16 // 保留97
	Reserved98 uint16 // 保留98
	Reserved99 uint16 // 保留99
	Reserved100 uint16 // 保留100
}

// DocTable .doc格式的表格
type DocTable struct {
	Rows    []*DocTableRow
	Style   string
	Width   int
	Height  int
}

// DocTableRow .doc格式的表格行
type DocTableRow struct {
	Cells []*DocTableCell
}

// DocTableCell .doc格式的表格单元格
type DocTableCell struct {
	Text  string
	Width int
}

// MacroContainer 宏容器
type MacroContainer struct {
	VBAProject *VBAProject
	Modules    []*VBAModule
}

// VBAProject VBA项目
type VBAProject struct {
	Name        string
	Description string
	Version     string
}

// VBAModule VBA模块
type VBAModule struct {
	Name     string
	Type     string
	Code     string
	Language string
}

 

// DocParagraph .doc格式的段落
type DocParagraph struct {
	Text      string
	Style     string
	Alignment string
	Indent    int
	Spacing   float64
}

// DocImage .doc格式的图片
type DocImage struct {
	ID      string
	Path    string
	Width   int
	Height  int
	AltText string
	Title   string
}

// DocStyle .doc格式的样式
type DocStyle struct {
	Name        string
	Type        string
	FontName    string
	FontSize    int
	Bold        bool
	Italic      bool
	Underline   bool
	Alignment   string
	Indent      int
	Spacing     float64
}

// ProgressCallback 进度回调函数类型
type ProgressCallback func(progress float64, message string) 