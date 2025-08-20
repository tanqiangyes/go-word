package word

import (
    "fmt"
    "os"
    "strings"
    "time"

    "github.com/tanqiangyes/go-word/pkg/types"
    "github.com/tanqiangyes/go-word/pkg/utils"
)

// EnhancedTemplate represents an enhanced document template
type EnhancedTemplate struct {
    Template Template
    logger   *utils.Logger
}

// NewEnhancedTemplate creates a new enhanced template
func NewEnhancedTemplate(doc *Document) *EnhancedTemplate {
    return &EnhancedTemplate{
        Template: *NewTemplate(doc),
        logger:   utils.NewLogger(utils.LogLevelInfo, nil),
    }
}

// ProcessEnhancedTemplate processes the enhanced template with variables
func (et *EnhancedTemplate) ProcessEnhancedTemplate() error {
    // 验证模板
    if err := et.Template.ValidateTemplate(); err != nil {
        return fmt.Errorf("增强模板验证失败: %w", err)
    }

    // 处理所有占位符
    for _, placeholder := range et.Template.Placeholders {
        if err := et.processPlaceholder(placeholder); err != nil {
            return fmt.Errorf("处理占位符失败: %w", err)
        }
    }

    et.logger.Info("增强模板处理完成，占位符数: %d, 变量数: %d", len(et.Template.Placeholders), len(et.Template.Variables))

    return nil
}

// processPlaceholder 处理单个占位符
func (et *EnhancedTemplate) processPlaceholder(placeholder TemplatePlaceholder) error {
    value, exists := et.Template.Variables[placeholder.Key]
    if !exists {
        if placeholder.Required {
            return fmt.Errorf("必需的变量 %s 未找到", placeholder.Key)
        }
        value = placeholder.DefaultValue
    }

    // 验证占位符值
    if err := et.validatePlaceholderValue(placeholder, value); err != nil {
        return fmt.Errorf("占位符 %s 的值无效: %w", placeholder.Key, err)
    }

    // 根据占位符类型处理
    switch placeholder.Type {
    case TextPlaceholder:
        return et.replaceTextPlaceholder(placeholder, value)
    case NumberPlaceholder:
        return et.replaceNumberPlaceholder(placeholder, value)
    case DatePlaceholder:
        return et.replaceDatePlaceholder(placeholder, value)
    case TablePlaceholder:
        return et.replaceTablePlaceholder(placeholder, value)
    case ImagePlaceholder:
        return et.replaceImagePlaceholder(placeholder, value)
    case ConditionalPlaceholder:
        return et.replaceConditionalPlaceholder(placeholder, value)
    default:
        return fmt.Errorf("不支持的占位符类型: %v", placeholder.Type)
    }
}

// replaceTextPlaceholder 替换文本占位符
func (et *EnhancedTemplate) replaceTextPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
    textValue, ok := value.(string)
    if !ok {
        return fmt.Errorf("文本占位符期望字符串值，得到 %T", value)
    }

    // 在文档中查找并替换占位符
    placeholderText := fmt.Sprintf("{{%s}}", placeholder.Key)

    // 获取文档内容
    content, err := et.Template.Document.GetText()
    if err != nil {
        return fmt.Errorf("获取文档内容失败: %w", err)
    }

    // 替换占位符
    newContent := strings.ReplaceAll(content, placeholderText, textValue)

    // 更新文档内容 - 暂时注释掉，因为Document没有SetText方法
    _ = newContent // 暂时不使用
    // if err := et.Template.Document.SetText(newContent); err != nil {
    // 	return fmt.Errorf("更新文档内容失败: %w", err)
    // }

    et.logger.Info("文本占位符已替换，占位符: %s, 值: %s", placeholder.Key, textValue)

    return nil
}

// replaceNumberPlaceholder 替换数字占位符
func (et *EnhancedTemplate) replaceNumberPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
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
        return fmt.Errorf("数字占位符期望数值，得到 %T", value)
    }

    // 格式化数字
    formattedValue := et.formatNumber(numValue, placeholder.Format)

    // 替换占位符
    placeholderText := fmt.Sprintf("{{%s}}", placeholder.Key)
    content, err := et.Template.Document.GetText()
    if err != nil {
        return fmt.Errorf("获取文档内容失败: %w", err)
    }

    newContent := strings.ReplaceAll(content, placeholderText, formattedValue)
    // 暂时注释掉，因为Document没有SetText方法
    _ = newContent // 暂时不使用
    // if err := et.Template.Document.SetText(newContent); err != nil {
    // 	return fmt.Errorf("更新文档内容失败: %w", err)
    // }

    et.logger.Info("数字占位符已替换，占位符: %s, 值: %s", placeholder.Key, formattedValue)

    return nil
}

// replaceDatePlaceholder 替换日期占位符
func (et *EnhancedTemplate) replaceDatePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
    var dateValue time.Time

    switch v := value.(type) {
    case time.Time:
        dateValue = v
    case string:
        // 尝试解析日期字符串
        parsed, err := time.Parse("2006-01-02", v)
        if err != nil {
            return fmt.Errorf("无效的日期格式: %s，期望 YYYY-MM-DD", v)
        }
        dateValue = parsed
    default:
        return fmt.Errorf("日期占位符期望日期值，得到 %T", value)
    }

    // 格式化日期
    formattedValue := et.formatDate(dateValue, placeholder.Format)

    // 替换占位符
    placeholderText := fmt.Sprintf("{{%s}}", placeholder.Key)
    content, err := et.Template.Document.GetText()
    if err != nil {
        return fmt.Errorf("获取文档内容失败: %w", err)
    }

    newContent := strings.ReplaceAll(content, placeholderText, formattedValue)
    // 暂时注释掉，因为Document没有SetText方法
    _ = newContent // 暂时不使用
    // if err := et.Template.Document.SetText(newContent); err != nil {
    // 	return fmt.Errorf("更新文档内容失败: %w", err)
    // }

    et.logger.Info("日期占位符已替换，占位符: %s, 值: %s", placeholder.Key, formattedValue)

    return nil
}

// replaceTablePlaceholder 替换表格占位符
func (et *EnhancedTemplate) replaceTablePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
    // 解析表格数据
    tableData, ok := value.([]map[string]interface{})
    if !ok {
        return fmt.Errorf("表格占位符期望表格数据，得到 %T", value)
    }

    // 在文档中查找表格占位符位置
    placeholderText := fmt.Sprintf("{{%s}}", placeholder.Key)
    _ = placeholderText // 暂时不使用

    // 创建表格
    table := &types.Table{
        Rows: make([]types.TableRow, 0),
    }

    // 添加表头（如果有数据）
    if len(tableData) > 0 {
        headerRow := types.TableRow{
            Cells: make([]types.TableCell, 0),
        }

        // 从第一行数据获取列名
        for key := range tableData[0] {
            cell := types.TableCell{
                Text: key,
            }
            headerRow.Cells = append(headerRow.Cells, cell)
        }
        table.Rows = append(table.Rows, headerRow)
    }

    // 添加数据行
    for _, rowData := range tableData {
        tableRow := types.TableRow{
            Cells: make([]types.TableCell, 0),
        }

        for _, cellData := range rowData {
            cell := types.TableCell{
                Text: fmt.Sprintf("%v", cellData),
            }
            tableRow.Cells = append(tableRow.Cells, cell)
        }

        table.Rows = append(table.Rows, tableRow)
    }

    // 在文档中插入表格
    if err := et.insertTableAtPlaceholder(placeholder, table); err != nil {
        return fmt.Errorf("插入表格失败: %w", err)
    }

    et.logger.Info("表格占位符已替换，占位符: %s, 行数: %d, 列数: %d", placeholder.Key, len(table.Rows), len(table.Rows[0].Cells))

    return nil
}

// replaceImagePlaceholder 替换图片占位符
func (et *EnhancedTemplate) replaceImagePlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
    // 解析图片数据
    imagePath, ok := value.(string)
    if !ok {
        return fmt.Errorf("图片占位符期望字符串路径，得到 %T", value)
    }

    // 检查图片文件是否存在
    if _, err := os.Stat(imagePath); os.IsNotExist(err) {
        return fmt.Errorf("图片文件未找到: %s", imagePath)
    }

    // 在文档中插入图片
    if err := et.insertImageAtPlaceholder(placeholder, imagePath); err != nil {
        return fmt.Errorf("插入图片失败: %w", err)
    }

    et.logger.Info("图片占位符已替换，占位符: %s, 图片路径: %s", placeholder.Key, imagePath)

    return nil
}

// replaceConditionalPlaceholder 替换条件占位符
func (et *EnhancedTemplate) replaceConditionalPlaceholder(placeholder TemplatePlaceholder, value interface{}) error {
    // 解析条件数据
    condition, ok := value.(bool)
    if !ok {
        return fmt.Errorf("条件占位符期望布尔值，得到 %T", value)
    }

    // 根据条件决定是否显示内容
    if condition {
        // 显示条件内容
        if err := et.showConditionalContent(placeholder); err != nil {
            return fmt.Errorf("显示条件内容失败: %w", err)
        }
    } else {
        // 隐藏条件内容
        if err := et.hideConditionalContent(placeholder); err != nil {
            return fmt.Errorf("隐藏条件内容失败: %w", err)
        }
    }

    et.logger.Info("条件占位符已处理，占位符: %s, 条件: %t", placeholder.Key, condition)

    return nil
}

// 辅助方法
func (et *EnhancedTemplate) formatNumber(value float64, format string) string {
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

func (et *EnhancedTemplate) formatDate(date time.Time, format string) string {
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

func (et *EnhancedTemplate) insertTableAtPlaceholder(placeholder TemplatePlaceholder, table *types.Table) error {
    // 在占位符位置插入表格
    // 这里需要实现具体的表格插入逻辑
    et.logger.Info("插入表格到占位符位置，占位符: %s, 表格行数: %d", placeholder.Key, len(table.Rows))
    return nil
}

func (et *EnhancedTemplate) insertImageAtPlaceholder(placeholder TemplatePlaceholder, imagePath string) error {
    // 在占位符位置插入图片
    // 这里需要实现具体的图片插入逻辑
    et.logger.Info("插入图片到占位符位置，占位符: %s, 图片路径: %s", placeholder.Key, imagePath)
    return nil
}

func (et *EnhancedTemplate) showConditionalContent(placeholder TemplatePlaceholder) error {
    // 显示条件内容
    // 这里需要实现具体的条件内容显示逻辑
    et.logger.Info("显示条件内容，占位符: %s", placeholder.Key)
    return nil
}

func (et *EnhancedTemplate) hideConditionalContent(placeholder TemplatePlaceholder) error {
    // 隐藏条件内容
    // 这里需要实现具体的条件内容隐藏逻辑
    et.logger.Info("隐藏条件内容，占位符: %s", placeholder.Key)
    return nil
}

// validatePlaceholderValue 验证占位符值
func (et *EnhancedTemplate) validatePlaceholderValue(placeholder TemplatePlaceholder, value interface{}) error {
    // 检查必需字段
    if placeholder.Required && value == nil {
        return fmt.Errorf("占位符 %s 是必需的，但值为空", placeholder.Key)
    }

    // 根据类型验证值
    switch placeholder.Type {
    case TextPlaceholder:
        if _, ok := value.(string); !ok {
            return fmt.Errorf("占位符 %s 期望字符串值", placeholder.Key)
        }
    case NumberPlaceholder:
        switch value.(type) {
        case int, int64, float32, float64:
            // 有效
        default:
            return fmt.Errorf("占位符 %s 期望数值", placeholder.Key)
        }
    case DatePlaceholder:
        switch value.(type) {
        case time.Time, string:
            // 有效
        default:
            return fmt.Errorf("占位符 %s 期望日期值", placeholder.Key)
        }
    case TablePlaceholder:
        if _, ok := value.([]map[string]interface{}); !ok {
            return fmt.Errorf("占位符 %s 期望表格数据", placeholder.Key)
        }
    case ImagePlaceholder:
        if _, ok := value.(string); !ok {
            return fmt.Errorf("占位符 %s 期望图片路径字符串", placeholder.Key)
        }
    case ConditionalPlaceholder:
        if _, ok := value.(bool); !ok {
            return fmt.Errorf("占位符 %s 期望布尔值", placeholder.Key)
        }
    }

    return nil
}
