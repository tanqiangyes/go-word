// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"regexp"
	"strings"
	"time"

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

	// 验证占位符值
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

// replaceTextPlaceholder 替换文本占位符
func (t *Template) replaceTextPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	textValue, ok := value.(string)
	if !ok {
		return fmt.Errorf("expected string value for text placeholder, got %T", value)
	}

	// 在文档中查找并替换占位符
	placeholderText := fmt.Sprintf("{{%s}}", placeholder.Key)
	
	// 获取文档内容
	content, err := t.Document.GetText()
	if err != nil {
		return fmt.Errorf("failed to get document content: %w", err)
	}

	// 替换占位符
	newContent := strings.ReplaceAll(content, placeholderText, textValue)
	
	// 更新文档内容
	if err := t.Document.SetText(newContent); err != nil {
		return fmt.Errorf("failed to update document content: %w", err)
	}

	t.logger.Info("文本占位符已替换", map[string]interface{}{
		"placeholder": placeholder.Key,
		"value":       textValue,
	})

	return nil
}

// replaceNumberPlaceholder 替换数字占位符
func (t *Template) replaceNumberPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	var numValue float64
	
	switch v := value.(type) {
	case int:
		numValue = float64(v)
	case int64:
		numValue = float64(v)
	case float32:
		numValue = float64(v)
	case float64:
		numValue = v
	default:
		return fmt.Errorf("expected numeric value for number placeholder, got %T", value)
	}

	// 格式化数字
	formattedValue := t.formatNumber(numValue, placeholder.Format)
	
	// 替换占位符
	placeholderText := fmt.Sprintf("{{%s}}", placeholder.Key)
	content, err := t.Document.GetText()
	if err != nil {
		return fmt.Errorf("failed to get document content: %w", err)
	}

	newContent := strings.ReplaceAll(content, placeholderText, formattedValue)
	if err := t.Document.SetText(newContent); err != nil {
		return fmt.Errorf("failed to update document content: %w", err)
	}

	t.logger.Info("数字占位符已替换", map[string]interface{}{
		"placeholder": placeholder.Key,
		"value":       formattedValue,
	})

	return nil
}

// replaceDatePlaceholder 替换日期占位符
func (t *Template) replaceDatePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	var dateValue time.Time
	
	switch v := value.(type) {
	case time.Time:
		dateValue = v
	case string:
		// 尝试解析日期字符串
		parsed, err := time.Parse("2006-01-02", v)
		if err != nil {
			return fmt.Errorf("invalid date format: %s, expected YYYY-MM-DD", v)
		}
		dateValue = parsed
	default:
		return fmt.Errorf("expected date value for date placeholder, got %T", value)
	}

	// 格式化日期
	formattedValue := t.formatDate(dateValue, placeholder.Format)
	
	// 替换占位符
	placeholderText := fmt.Sprintf("{{%s}}", placeholder.Key)
	content, err := t.Document.GetText()
	if err != nil {
		return fmt.Errorf("failed to get document content: %w", err)
	}

	newContent := strings.ReplaceAll(content, placeholderText, formattedValue)
	if err := t.Document.SetText(newContent); err != nil {
		return fmt.Errorf("failed to update document content: %w", err)
	}

	t.logger.Info("日期占位符已替换", map[string]interface{}{
		"placeholder": placeholder.Key,
		"value":       formattedValue,
	})

	return nil
}

// replaceTablePlaceholder 替换表格占位符
func (t *Template) replaceTablePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	// 解析表格数据
	tableData, ok := value.([]map[string]interface{})
	if !ok {
		return fmt.Errorf("expected table data for table placeholder, got %T", value)
	}

	// 在文档中查找表格占位符位置
	placeholderText := fmt.Sprintf("{{%s}}", placeholder.Key)
	
	// 创建表格
	table := &types.Table{
		Rows: make([]*types.TableRow, 0),
	}

	// 添加表头（如果有数据）
	if len(tableData) > 0 {
		headerRow := &types.TableRow{
			Cells: make([]*types.TableCell, 0),
		}
		
		// 从第一行数据获取列名
		for key := range tableData[0] {
			cell := &types.TableCell{
				Text: key,
			}
			headerRow.Cells = append(headerRow.Cells, cell)
		}
		table.Rows = append(table.Rows, headerRow)
	}

	// 添加数据行
	for _, rowData := range tableData {
		tableRow := &types.TableRow{
			Cells: make([]*types.TableCell, 0),
		}
		
		for _, cellData := range rowData {
			cell := &types.TableCell{
				Text: fmt.Sprintf("%v", cellData),
			}
			tableRow.Cells = append(tableRow.Cells, cell)
		}
		
		table.Rows = append(table.Rows, tableRow)
	}

	// 在文档中插入表格
	if err := t.insertTableAtPlaceholder(placeholder, table); err != nil {
		return fmt.Errorf("failed to insert table: %w", err)
	}

	t.logger.Info("表格占位符已替换", map[string]interface{}{
		"placeholder": placeholder.Key,
		"rows":        len(table.Rows),
		"columns":     len(table.Rows[0].Cells),
	})

	return nil
}

// replaceImagePlaceholder 替换图片占位符
func (t *Template) replaceImagePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	// 解析图片数据
	imagePath, ok := value.(string)
	if !ok {
		return fmt.Errorf("expected string path for image placeholder, got %T", value)
	}

	// 检查图片文件是否存在
	if _, err := os.Stat(imagePath); os.IsNotExist(err) {
		return fmt.Errorf("image file not found: %s", imagePath)
	}

	// 在文档中插入图片
	if err := t.insertImageAtPlaceholder(placeholder, imagePath); err != nil {
		return fmt.Errorf("failed to insert image: %w", err)
	}

	t.logger.Info("图片占位符已替换", map[string]interface{}{
		"placeholder": placeholder.Key,
		"image_path":  imagePath,
	})

	return nil
}

// replaceConditionalPlaceholder 替换条件占位符
func (t *Template) replaceConditionalPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
	// 解析条件数据
	condition, ok := value.(bool)
	if !ok {
		return fmt.Errorf("expected boolean value for conditional placeholder, got %T", value)
	}

	// 根据条件决定是否显示内容
	if condition {
		// 显示条件内容
		if err := t.showConditionalContent(placeholder); err != nil {
			return fmt.Errorf("failed to show conditional content: %w", err)
		}
	} else {
		// 隐藏条件内容
		if err := t.hideConditionalContent(placeholder); err != nil {
			return fmt.Errorf("failed to hide conditional content: %w", err)
		}
	}

	t.logger.Info("条件占位符已处理", map[string]interface{}{
		"placeholder": placeholder.Key,
		"condition":   condition,
	})

	return nil
}

// 辅助方法
func (t *Template) formatNumber(value float64, format string) string {
	if format == "" {
		return fmt.Sprintf("%.2f", value)
	}
	
	switch format {
	case "integer":
		return fmt.Sprintf("%.0f", value)
	case "currency":
		return fmt.Sprintf("¥%.2f", value)
	case "percentage":
		return fmt.Sprintf("%.1f%%", value*100)
	default:
		return fmt.Sprintf("%.2f", value)
	}
}

func (t *Template) formatDate(date time.Time, format string) string {
	if format == "" {
		return date.Format("2006-01-02")
	}
	
	switch format {
	case "short":
		return date.Format("01/02/06")
	case "long":
		return date.Format("January 2, 2006")
	case "time":
		return date.Format("15:04:05")
	case "datetime":
		return date.Format("2006-01-02 15:04:05")
	default:
		return date.Format("2006-01-02")
	}
}

func (t *Template) insertTableAtPlaceholder(placeholder TemplatePlaceholder, table *types.Table) error {
	// 在占位符位置插入表格
	// 这里需要实现具体的表格插入逻辑
	return nil
}

func (t *Template) insertImageAtPlaceholder(placeholder TemplatePlaceholder, imagePath string) error {
	// 在占位符位置插入图片
	// 这里需要实现具体的图片插入逻辑
	return nil
}

func (t *Template) showConditionalContent(placeholder TemplatePlaceholder) error {
	// 显示条件内容
	// 这里需要实现具体的条件内容显示逻辑
	return nil
}

func (t *Template) hideConditionalContent(placeholder TemplatePlaceholder) error {
	// 隐藏条件内容
	// 这里需要实现具体的条件内容隐藏逻辑
	return nil
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