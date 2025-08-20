// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// DocumentValidator represents a document validator
type DocumentValidator struct {
	Document     *Document
	ValidationRules []ValidationRule
	ValidationResults []ValidationResult
	AutoFix       bool
}

// ValidationRule represents a validation rule
type ValidationRule struct {
	ID          string
	Name        string
	Description string
	Severity    ValidationSeverity
	Validator   func(*Document) ValidationResult
	AutoFixer   func(*Document) error
}

// ValidationSeverity 使用 advanced_tables.go 中的定义

// ValidationResult represents a validation result
type ValidationResult struct {
	RuleID      string
	Severity    ValidationSeverity
	Message     string
	Location    string
	Fixed       bool
	Error       error
}

// NewDocumentValidator creates a new document validator
func NewDocumentValidator(doc *Document) *DocumentValidator {
	validator := &DocumentValidator{
		Document:        doc,
		ValidationRules: make([]ValidationRule, 0),
		ValidationResults: make([]ValidationResult, 0),
		AutoFix:        false,
	}

	// 添加默认验证规则
	validator.addDefaultRules()

	return validator
}

// addDefaultRules adds default validation rules
func (dv *DocumentValidator) addDefaultRules() {
	// 文档完整性验证
	dv.AddRule(ValidationRule{
		ID:          "doc_integrity",
		Name:        "文档完整性",
		Description: "检查文档是否完整且可读",
		Severity:    CriticalSeverity,
		Validator:   dv.validateDocumentIntegrity,
		AutoFixer:   dv.fixDocumentIntegrity,
	})

	// 内容结构验证
	dv.AddRule(ValidationRule{
		ID:          "content_structure",
		Name:        "内容结构",
		Description: "检查文档内容结构是否合理",
		Severity:    WarningSeverity,
		Validator:   dv.validateContentStructure,
		AutoFixer:   dv.fixContentStructure,
	})

	// 格式规范验证
	dv.AddRule(ValidationRule{
		ID:          "format_standards",
		Name:        "格式规范",
		Description: "检查文档格式是否符合规范",
		Severity:    WarningSeverity,
		Validator:   dv.validateFormatStandards,
		AutoFixer:   dv.fixFormatStandards,
	})

	// 文本质量验证
	dv.AddRule(ValidationRule{
		ID:          "text_quality",
		Name:        "文本质量",
		Description: "检查文本质量和一致性",
		Severity:    InfoSeverity,
		Validator:   dv.validateTextQuality,
		AutoFixer:   dv.fixTextQuality,
	})

	// 表格结构验证
	dv.AddRule(ValidationRule{
		ID:          "table_structure",
		Name:        "表格结构",
		Description: "检查表格结构是否合理",
		Severity:    WarningSeverity,
		Validator:   dv.validateTableStructure,
		AutoFixer:   dv.fixTableStructure,
	})
}

// AddRule adds a validation rule
func (dv *DocumentValidator) AddRule(rule ValidationRule) {
	dv.ValidationRules = append(dv.ValidationRules, rule)
}

// ValidateDocument validates the document
func (dv *DocumentValidator) ValidateDocument() error {
	if dv.Document == nil {
		return fmt.Errorf("document is nil")
	}

	dv.ValidationResults = make([]ValidationResult, 0)

	// 执行所有验证规则
	for _, rule := range dv.ValidationRules {
		result := rule.Validator(dv.Document)
		result.RuleID = rule.ID
		
		dv.ValidationResults = append(dv.ValidationResults, result)

		// 如果启用了自动修复且验证失败
		if dv.AutoFix && result.Error != nil && rule.AutoFixer != nil {
			if err := rule.AutoFixer(dv.Document); err == nil {
				result.Fixed = true
				result.Message = "已自动修复: " + result.Message
			}
		}
	}

	return nil
}

// validateDocumentIntegrity validates document integrity
func (dv *DocumentValidator) validateDocumentIntegrity(doc *Document) ValidationResult {
	if doc.mainPart == nil {
		return ValidationResult{
			Severity: CriticalSeverity,
			Message:  "文档缺少主部分",
			Location: "document",
			Error:    fmt.Errorf("main part is nil"),
		}
	}

	if doc.mainPart.Content == nil {
		return ValidationResult{
			Severity: CriticalSeverity,
			Message:  "文档缺少内容",
			Location: "document.content",
			Error:    fmt.Errorf("content is nil"),
		}
	}

	if len(doc.mainPart.Content.Paragraphs) == 0 {
		return ValidationResult{
			Severity: WarningSeverity,
			Message:  "文档没有段落内容",
			Location: "document.content.paragraphs",
			Error:    fmt.Errorf("no paragraphs found"),
		}
	}

	return ValidationResult{
		Severity: InfoSeverity,
		Message:  "文档完整性验证通过",
		Location: "document",
	}
}

// fixDocumentIntegrity fixes document integrity issues
func (dv *DocumentValidator) fixDocumentIntegrity(doc *Document) error {
	// 如果主部分为空，创建默认内容
	if doc.mainPart == nil {
		doc.mainPart = &MainDocumentPart{
			Content: &types.DocumentContent{
				Paragraphs: make([]types.Paragraph, 0),
				Tables:     make([]types.Table, 0),
				Text:       "",
			},
		}
	}

	if doc.mainPart.Content == nil {
		doc.mainPart.Content = &types.DocumentContent{
			Paragraphs: make([]types.Paragraph, 0),
			Tables:     make([]types.Table, 0),
			Text:       "",
		}
	}

	return nil
}

// validateContentStructure validates content structure
func (dv *DocumentValidator) validateContentStructure(doc *Document) ValidationResult {
	if doc.mainPart == nil || doc.mainPart.Content == nil {
		return ValidationResult{
			Severity: CriticalSeverity,
			Message:  "无法验证内容结构：文档内容为空",
			Location: "document.content",
			Error:    fmt.Errorf("content is nil"),
		}
	}

	content := doc.mainPart.Content
	issues := make([]string, 0)

	// 检查段落结构
	for i, paragraph := range content.Paragraphs {
		if strings.TrimSpace(paragraph.Text) == "" && len(paragraph.Runs) == 0 {
			issues = append(issues, fmt.Sprintf("段落 %d 为空", i+1))
		}
	}

	// 检查表格结构
	for i, table := range content.Tables {
		if len(table.Rows) == 0 {
			issues = append(issues, fmt.Sprintf("表格 %d 没有行", i+1))
		}
		for j, row := range table.Rows {
			if len(row.Cells) == 0 {
				issues = append(issues, fmt.Sprintf("表格 %d 行 %d 没有单元格", i+1, j+1))
			}
		}
	}

	if len(issues) > 0 {
		return ValidationResult{
			Severity: WarningSeverity,
			Message:  "内容结构问题: " + strings.Join(issues, "; "),
			Location: "document.content",
			Error:    fmt.Errorf("content structure issues found"),
		}
	}

	return ValidationResult{
		Severity: InfoSeverity,
		Message:  "内容结构验证通过",
		Location: "document.content",
	}
}

// fixContentStructure fixes content structure issues
func (dv *DocumentValidator) fixContentStructure(doc *Document) error {
	if doc.mainPart == nil || doc.mainPart.Content == nil {
		return fmt.Errorf("cannot fix content structure: content is nil")
	}

	content := doc.mainPart.Content

	// 移除空段落
	validParagraphs := make([]types.Paragraph, 0)
	for _, paragraph := range content.Paragraphs {
		if strings.TrimSpace(paragraph.Text) != "" || len(paragraph.Runs) > 0 {
			validParagraphs = append(validParagraphs, paragraph)
		}
	}
	content.Paragraphs = validParagraphs

	// 移除空表格
	validTables := make([]types.Table, 0)
	for _, table := range content.Tables {
		if len(table.Rows) > 0 {
			validTables = append(validTables, table)
		}
	}
	content.Tables = validTables

	return nil
}

// validateFormatStandards validates format standards
func (dv *DocumentValidator) validateFormatStandards(doc *Document) ValidationResult {
	if doc.mainPart == nil || doc.mainPart.Content == nil {
		return ValidationResult{
			Severity: CriticalSeverity,
			Message:  "无法验证格式规范：文档内容为空",
			Location: "document.content",
			Error:    fmt.Errorf("content is nil"),
		}
	}

	content := doc.mainPart.Content
	issues := make([]string, 0)

	// 检查段落格式
	for i, paragraph := range content.Paragraphs {
		for j, run := range paragraph.Runs {
			// 检查字体大小
			if run.FontSize <= 0 {
				issues = append(issues, fmt.Sprintf("段落 %d 运行 %d 字体大小无效", i+1, j+1))
			}

			// 检查字体名称
			if run.FontName == "" {
				issues = append(issues, fmt.Sprintf("段落 %d 运行 %d 字体名称为空", i+1, j+1))
			}
		}
	}

	if len(issues) > 0 {
		return ValidationResult{
			Severity: WarningSeverity,
			Message:  "格式规范问题: " + strings.Join(issues, "; "),
			Location: "document.content",
			Error:    fmt.Errorf("format standard issues found"),
		}
	}

	return ValidationResult{
		Severity: InfoSeverity,
		Message:  "格式规范验证通过",
		Location: "document.content",
	}
}

// fixFormatStandards fixes format standard issues
func (dv *DocumentValidator) fixFormatStandards(doc *Document) error {
	if doc.mainPart == nil || doc.mainPart.Content == nil {
		return fmt.Errorf("cannot fix format standards: content is nil")
	}

	content := doc.mainPart.Content

	// 修复段落格式
	for i := range content.Paragraphs {
		for j := range content.Paragraphs[i].Runs {
			run := &content.Paragraphs[i].Runs[j]
			
			// 设置默认字体大小
			if run.FontSize <= 0 {
				run.FontSize = 12
			}

			// 设置默认字体名称
			if run.FontName == "" {
				run.FontName = "Arial"
			}
		}
	}

	return nil
}

// validateTextQuality validates text quality
func (dv *DocumentValidator) validateTextQuality(doc *Document) ValidationResult {
	if doc.mainPart == nil || doc.mainPart.Content == nil {
		return ValidationResult{
			Severity: CriticalSeverity,
			Message:  "无法验证文本质量：文档内容为空",
			Location: "document.content",
			Error:    fmt.Errorf("content is nil"),
		}
	}

	content := doc.mainPart.Content
	issues := make([]string, 0)

	// 检查文本质量
	for i, paragraph := range content.Paragraphs {
		text := strings.TrimSpace(paragraph.Text)
		
		// 检查重复空格
		if strings.Contains(text, "  ") {
			issues = append(issues, fmt.Sprintf("段落 %d 包含重复空格", i+1))
		}

		// 检查特殊字符
		if strings.ContainsAny(text, "") {
			issues = append(issues, fmt.Sprintf("段落 %d 包含特殊字符", i+1))
		}
	}

	if len(issues) > 0 {
		return ValidationResult{
			Severity: InfoSeverity,
			Message:  "文本质量问题: " + strings.Join(issues, "; "),
			Location: "document.content",
			Error:    fmt.Errorf("text quality issues found"),
		}
	}

	return ValidationResult{
		Severity: InfoSeverity,
		Message:  "文本质量验证通过",
		Location: "document.content",
	}
}

// fixTextQuality fixes text quality issues
func (dv *DocumentValidator) fixTextQuality(doc *Document) error {
	if doc.mainPart == nil || doc.mainPart.Content == nil {
		return fmt.Errorf("cannot fix text quality: content is nil")
	}

	content := doc.mainPart.Content

	// 修复文本质量
	for i := range content.Paragraphs {
		// 修复重复空格
		content.Paragraphs[i].Text = strings.Join(strings.Fields(content.Paragraphs[i].Text), " ")
		
		// 修复运行中的文本
		for j := range content.Paragraphs[i].Runs {
			content.Paragraphs[i].Runs[j].Text = strings.Join(strings.Fields(content.Paragraphs[i].Runs[j].Text), " ")
		}
	}

	return nil
}

// validateTableStructure validates table structure
func (dv *DocumentValidator) validateTableStructure(doc *Document) ValidationResult {
	if doc.mainPart == nil || doc.mainPart.Content == nil {
		return ValidationResult{
			Severity: CriticalSeverity,
			Message:  "无法验证表格结构：文档内容为空",
			Location: "document.content",
			Error:    fmt.Errorf("content is nil"),
		}
	}

	content := doc.mainPart.Content
	issues := make([]string, 0)

	// 检查表格结构
	for i, table := range content.Tables {
		if table.Columns <= 0 {
			issues = append(issues, fmt.Sprintf("表格 %d 列数无效", i+1))
		}

		for j, row := range table.Rows {
			if len(row.Cells) != table.Columns {
				issues = append(issues, fmt.Sprintf("表格 %d 行 %d 单元格数量不匹配", i+1, j+1))
			}
		}
	}

	if len(issues) > 0 {
		return ValidationResult{
			Severity: WarningSeverity,
			Message:  "表格结构问题: " + strings.Join(issues, "; "),
			Location: "document.content.tables",
			Error:    fmt.Errorf("table structure issues found"),
		}
	}

	return ValidationResult{
		Severity: InfoSeverity,
		Message:  "表格结构验证通过",
		Location: "document.content.tables",
	}
}

// fixTableStructure fixes table structure issues
func (dv *DocumentValidator) fixTableStructure(doc *Document) error {
	if doc.mainPart == nil || doc.mainPart.Content == nil {
		return fmt.Errorf("cannot fix table structure: content is nil")
	}

	content := doc.mainPart.Content

	// 修复表格结构
	for i := range content.Tables {
		table := &content.Tables[i]
		
		// 计算正确的列数
		maxColumns := 0
		for _, row := range table.Rows {
			if len(row.Cells) > maxColumns {
				maxColumns = len(row.Cells)
			}
		}
		
		table.Columns = maxColumns

		// 确保每行的单元格数量一致
		for j := range table.Rows {
			row := &table.Rows[j]
			for len(row.Cells) < maxColumns {
				row.Cells = append(row.Cells, types.TableCell{Text: ""})
			}
		}
	}

	return nil
}

// GetValidationResults returns all validation results
func (dv *DocumentValidator) GetValidationResults() []ValidationResult {
	return dv.ValidationResults
}

// GetErrors returns only error-level validation results
func (dv *DocumentValidator) GetErrors() []ValidationResult {
	var errors []ValidationResult
	for _, result := range dv.ValidationResults {
		if result.Severity == ErrorSeverity || result.Severity == CriticalSeverity {
			errors = append(errors, result)
		}
	}
	return errors
}

// GetWarnings returns only warning-level validation results
func (dv *DocumentValidator) GetWarnings() []ValidationResult {
	var warnings []ValidationResult
	for _, result := range dv.ValidationResults {
		if result.Severity == WarningSeverity {
			warnings = append(warnings, result)
		}
	}
	return warnings
}

// GetInfo returns only info-level validation results
func (dv *DocumentValidator) GetInfo() []ValidationResult {
	var info []ValidationResult
	for _, result := range dv.ValidationResults {
		if result.Severity == InfoSeverity {
			info = append(info, result)
		}
	}
	return info
}

// HasErrors returns true if there are any errors
func (dv *DocumentValidator) HasErrors() bool {
	return len(dv.GetErrors()) > 0
}

// HasWarnings returns true if there are any warnings
func (dv *DocumentValidator) HasWarnings() bool {
	return len(dv.GetWarnings()) > 0
}

// GetValidationSummary returns a summary of validation results
func (dv *DocumentValidator) GetValidationSummary() string {
	var summary strings.Builder
	summary.WriteString("文档验证摘要:\n")
	summary.WriteString(fmt.Sprintf("验证规则数量: %d\n", len(dv.ValidationRules)))
	summary.WriteString(fmt.Sprintf("验证结果数量: %d\n", len(dv.ValidationResults)))
	summary.WriteString(fmt.Sprintf("错误数量: %d\n", len(dv.GetErrors())))
	summary.WriteString(fmt.Sprintf("警告数量: %d\n", len(dv.GetWarnings())))
	summary.WriteString(fmt.Sprintf("信息数量: %d\n", len(dv.GetInfo())))

	if dv.HasErrors() {
		summary.WriteString("\n错误详情:\n")
		for _, error := range dv.GetErrors() {
			summary.WriteString(fmt.Sprintf("- %s: %s\n", error.RuleID, error.Message))
		}
	}

	if dv.HasWarnings() {
		summary.WriteString("\n警告详情:\n")
		for _, warning := range dv.GetWarnings() {
			summary.WriteString(fmt.Sprintf("- %s: %s\n", warning.RuleID, warning.Message))
		}
	}

	return summary.String()
}

// SetAutoFix enables or disables auto-fix
func (dv *DocumentValidator) SetAutoFix(enabled bool) {
	dv.AutoFix = enabled
}

// String returns the string representation of ValidationSeverity
func (vs ValidationSeverity) String() string {
	switch vs {
	case InfoSeverity:
		return "Info"
	case WarningSeverity:
		return "Warning"
	case ErrorSeverity:
		return "Error"
	case CriticalSeverity:
		return "Critical"
	default:
		return "Unknown"
	}
} 