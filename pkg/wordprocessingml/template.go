// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// Template represents a document template
type Template struct {
	Document     *Document
	Variables    map[string]interface{}
	Placeholders []TemplatePlaceholder
	Validation   TemplateValidation
}

// TemplatePlaceholder represents a placeholder in the template
type TemplatePlaceholder struct {
	Key         string
	Type        PlaceholderType
	DefaultValue interface{}
	Required    bool
	Validation  string
	Position    PlaceholderPosition
}

// PlaceholderType defines the type of placeholder
type PlaceholderType int

const (
	// TextPlaceholder for text replacement
	TextPlaceholder PlaceholderType = iota
	// NumberPlaceholder for numeric values
	NumberPlaceholder
	// DatePlaceholder for date values
	DatePlaceholder
	// TablePlaceholder for table insertion
	TablePlaceholder
	// ImagePlaceholder for image insertion
	ImagePlaceholder
	// ConditionalPlaceholder for conditional content
	ConditionalPlaceholder
)

// PlaceholderPosition represents the position of a placeholder
type PlaceholderPosition struct {
	ParagraphIndex int
	RunIndex       int
	StartOffset    int
	EndOffset      int
}

// TemplateValidation represents template validation rules
type TemplateValidation struct {
	RequiredVariables []string
	VariableTypes     map[string]PlaceholderType
	CustomValidators  map[string]func(interface{}) error
}

// NewTemplate creates a new document template
func NewTemplate(doc *Document) *Template {
	return &Template{
		Document:     doc,
		Variables:    make(map[string]interface{}),
		Placeholders: make([]TemplatePlaceholder, 0),
		Validation: TemplateValidation{
			RequiredVariables: make([]string, 0),
			VariableTypes:     make(map[string]PlaceholderType),
			CustomValidators:  make(map[string]func(interface{}) error),
		},
	}
}

// AddVariable adds a variable to the template
func (t *Template) AddVariable(key string, value interface{}) {
	t.Variables[key] = value
}

// AddPlaceholder adds a placeholder to the template
func (t *Template) AddPlaceholder(placeholder TemplatePlaceholder) {
	t.Placeholders = append(t.Placeholders, placeholder)
	
	// 添加到验证规则
	if placeholder.Required {
		t.Validation.RequiredVariables = append(t.Validation.RequiredVariables, placeholder.Key)
	}
	
	if placeholder.Type != TextPlaceholder {
		t.Validation.VariableTypes[placeholder.Key] = placeholder.Type
	}
}

// ProcessTemplate processes the template with variables
func (t *Template) ProcessTemplate() error {
	// 验证模板
	if err := t.ValidateTemplate(); err != nil {
		return fmt.Errorf("template validation failed: %w", err)
	}

	// 处理所有占位符
	for _, placeholder := range t.Placeholders {
		if err := t.processPlaceholder(placeholder); err != nil {
			return fmt.Errorf("failed to process placeholder %s: %w", placeholder.Key, err)
		}
	}

	return nil
}

// processPlaceholder processes a single placeholder
func (t *Template) processPlaceholder(placeholder TemplatePlaceholder) error {
	value, exists := t.Variables[placeholder.Key]
	if !exists {
		if placeholder.Required {
			return fmt.Errorf("required variable %s not found", placeholder.Key)
		}
		value = placeholder.DefaultValue
	}

	// 验证值类型
	if err := t.validatePlaceholderValue(placeholder, value); err != nil {
		return fmt.Errorf("invalid value for placeholder %s: %w", placeholder.Key, err)
	}

	// 根据占位符类型处理
	switch placeholder.Type {
	case TextPlaceholder:
		return t.replaceTextPlaceholder(placeholder, value)
	case NumberPlaceholder:
		return t.replaceNumberPlaceholder(placeholder, value)
	case DatePlaceholder:
		return t.replaceDatePlaceholder(placeholder, value)
	case TablePlaceholder:
		return t.replaceTablePlaceholder(placeholder, value)
	case ImagePlaceholder:
		return t.replaceImagePlaceholder(placeholder, value)
	case ConditionalPlaceholder:
		return t.replaceConditionalPlaceholder(placeholder, value)
	default:
		return fmt.Errorf("unsupported placeholder type: %v", placeholder.Type)
	}
}

// replaceTextPlaceholder replaces a text placeholder
func (t *Template) replaceTextPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	textValue, ok := value.(string)
	if !ok {
		return fmt.Errorf("expected string value for text placeholder, got %T", value)
	}

	// 在文档中查找并替换占位符
	placeholderPattern := fmt.Sprintf("{{%s}}", placeholder.Key)
	replacementText := textValue

	return t.replaceInDocument(placeholderPattern, replacementText)
}

// replaceNumberPlaceholder replaces a number placeholder
func (t *Template) replaceNumberPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	// 将数字转换为字符串
	var numberText string
	switch v := value.(type) {
	case int:
		numberText = fmt.Sprintf("%d", v)
	case float64:
		numberText = fmt.Sprintf("%.2f", v)
	default:
		return fmt.Errorf("expected numeric value for number placeholder, got %T", value)
	}

	placeholderPattern := fmt.Sprintf("{{%s}}", placeholder.Key)
	return t.replaceInDocument(placeholderPattern, numberText)
}

// replaceDatePlaceholder replaces a date placeholder
func (t *Template) replaceDatePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	// 这里可以添加日期格式化逻辑
	dateText := fmt.Sprintf("%v", value)
	
	placeholderPattern := fmt.Sprintf("{{%s}}", placeholder.Key)
	return t.replaceInDocument(placeholderPattern, dateText)
}

// replaceTablePlaceholder replaces a table placeholder
func (t *Template) replaceTablePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	// 处理表格数据
	tableData, ok := value.([][]string)
	if !ok {
		return fmt.Errorf("expected table data for table placeholder, got %T", value)
	}

	// 创建表格
	table := types.Table{
		Rows:    make([]types.TableRow, 0, len(tableData)),
		Columns: len(tableData[0]),
	}

	for _, rowData := range tableData {
		row := types.TableRow{
			Cells: make([]types.TableCell, 0, len(rowData)),
		}
		
		for _, cellData := range rowData {
			cell := types.TableCell{
				Text: cellData,
			}
			row.Cells = append(row.Cells, cell)
		}
		
		table.Rows = append(table.Rows, row)
	}

	// 替换占位符为表格
	placeholderPattern := fmt.Sprintf("{{%s}}", placeholder.Key)
	return t.replaceTableInDocument(placeholderPattern, table)
}

// replaceImagePlaceholder replaces an image placeholder
func (t *Template) replaceImagePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	// 图片处理逻辑
	imagePath, ok := value.(string)
	if !ok {
		return fmt.Errorf("expected string value for image placeholder, got %T", value)
	}

	// 这里可以添加图片插入逻辑
	placeholderPattern := fmt.Sprintf("{{%s}}", placeholder.Key)
	return t.replaceInDocument(placeholderPattern, fmt.Sprintf("[图片: %s]", imagePath))
}

// replaceConditionalPlaceholder replaces a conditional placeholder
func (t *Template) replaceConditionalPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	// 条件内容处理
	condition, ok := value.(bool)
	if !ok {
		return fmt.Errorf("expected boolean value for conditional placeholder, got %T", value)
	}

	placeholderPattern := fmt.Sprintf("{{%s}}", placeholder.Key)
	if condition {
		return t.replaceInDocument(placeholderPattern, "是")
	} else {
		return t.replaceInDocument(placeholderPattern, "否")
	}
}

// replaceInDocument replaces text in the document
func (t *Template) replaceInDocument(pattern, replacement string) error {
	if t.Document.mainPart == nil || t.Document.mainPart.Content == nil {
		return fmt.Errorf("document content is nil")
	}

	content := t.Document.mainPart.Content
	
	// 替换段落中的文本
	for i := range content.Paragraphs {
		paragraph := &content.Paragraphs[i]
		
		// 替换段落文本
		paragraph.Text = strings.ReplaceAll(paragraph.Text, pattern, replacement)
		
		// 替换运行中的文本
		for j := range paragraph.Runs {
			run := &paragraph.Runs[j]
			run.Text = strings.ReplaceAll(run.Text, pattern, replacement)
		}
	}

	return nil
}

// replaceTableInDocument replaces a table placeholder with actual table
func (t *Template) replaceTableInDocument(pattern string, table types.Table) error {
	// 这里可以实现更复杂的表格替换逻辑
	// 暂时使用简单的文本替换
	tableText := "表格内容"
	return t.replaceInDocument(pattern, tableText)
}

// validatePlaceholderValue validates a placeholder value
func (t *Template) validatePlaceholderValue(placeholder TemplatePlaceholder, value interface{}) error {
	// 类型验证
	switch placeholder.Type {
	case TextPlaceholder:
		if _, ok := value.(string); !ok {
			return fmt.Errorf("expected string value for text placeholder")
		}
	case NumberPlaceholder:
		switch value.(type) {
		case int, float64:
			// 数字类型有效
		default:
			return fmt.Errorf("expected numeric value for number placeholder")
		}
	case DatePlaceholder:
		// 日期验证逻辑
	case TablePlaceholder:
		if _, ok := value.([][]string); !ok {
			return fmt.Errorf("expected table data for table placeholder")
		}
	case ImagePlaceholder:
		if _, ok := value.(string); !ok {
			return fmt.Errorf("expected string value for image placeholder")
		}
	case ConditionalPlaceholder:
		if _, ok := value.(bool); !ok {
			return fmt.Errorf("expected boolean value for conditional placeholder")
		}
	}

	// 自定义验证
	if validator, exists := t.Validation.CustomValidators[placeholder.Key]; exists {
		return validator(value)
	}

	return nil
}

// ValidateTemplate validates the template
func (t *Template) ValidateTemplate() error {
	// 检查必需变量
	for _, requiredVar := range t.Validation.RequiredVariables {
		if _, exists := t.Variables[requiredVar]; !exists {
			return fmt.Errorf("required variable %s not provided", requiredVar)
		}
	}

	// 检查变量类型
	for varName, expectedType := range t.Validation.VariableTypes {
		if value, exists := t.Variables[varName]; exists {
			if err := t.validateVariableType(varName, value, expectedType); err != nil {
				return err
			}
		}
	}

	return nil
}

// validateVariableType validates a variable type
func (t *Template) validateVariableType(varName string, value interface{}, expectedType PlaceholderType) error {
	switch expectedType {
	case TextPlaceholder:
		if _, ok := value.(string); !ok {
			return fmt.Errorf("variable %s should be string type", varName)
		}
	case NumberPlaceholder:
		switch value.(type) {
		case int, float64:
			// 有效
		default:
			return fmt.Errorf("variable %s should be numeric type", varName)
		}
	case DatePlaceholder:
		// 日期类型验证
	case TablePlaceholder:
		if _, ok := value.([][]string); !ok {
			return fmt.Errorf("variable %s should be table type", varName)
		}
	case ImagePlaceholder:
		if _, ok := value.(string); !ok {
			return fmt.Errorf("variable %s should be string type (image path)", varName)
		}
	case ConditionalPlaceholder:
		if _, ok := value.(bool); !ok {
			return fmt.Errorf("variable %s should be boolean type", varName)
		}
	}
	return nil
}

// ExtractPlaceholders extracts placeholders from the document
func (t *Template) ExtractPlaceholders() error {
	if t.Document.mainPart == nil || t.Document.mainPart.Content == nil {
		return fmt.Errorf("document content is nil")
	}

	content := t.Document.mainPart.Content
	placeholderRegex := regexp.MustCompile(`\{\{(\w+)\}\}`)

	for i, paragraph := range content.Paragraphs {
		matches := placeholderRegex.FindAllStringSubmatch(paragraph.Text, -1)
		
		for _, match := range matches {
			if len(match) >= 2 {
				placeholderKey := match[1]
				
				placeholder := TemplatePlaceholder{
					Key:         placeholderKey,
					Type:        TextPlaceholder, // 默认为文本类型
					Required:    true,
					Position: PlaceholderPosition{
						ParagraphIndex: i,
					},
				}
				
				t.AddPlaceholder(placeholder)
			}
		}
	}

	return nil
}

// GetTemplateSummary returns a summary of the template
func (t *Template) GetTemplateSummary() string {
	var summary strings.Builder
	summary.WriteString("模板摘要:\n")
	summary.WriteString(fmt.Sprintf("占位符数量: %d\n", len(t.Placeholders)))
	summary.WriteString(fmt.Sprintf("变量数量: %d\n", len(t.Variables)))
	summary.WriteString(fmt.Sprintf("必需变量: %d\n", len(t.Validation.RequiredVariables)))
	
	summary.WriteString("\n占位符列表:\n")
	for _, placeholder := range t.Placeholders {
		summary.WriteString(fmt.Sprintf("- %s (%v)\n", placeholder.Key, placeholder.Type))
	}
	
	return summary.String()
} 