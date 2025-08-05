package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewDocumentValidator(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	if validator == nil {
		t.Fatal("Expected DocumentValidator to be created")
	}
	
	if validator.Rules == nil {
		t.Error("Expected Rules to be initialized")
	}
	
	if validator.Results == nil {
		t.Error("Expected Results to be initialized")
	}
	
	if validator.AutoFix == false {
		t.Error("Expected AutoFix to be false by default")
	}
}

func TestAddRule(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 添加验证规则
	rule := wordprocessingml.ValidationRule{
		ID:          "test_rule",
		Name:        "Test Rule",
		Description: "A test validation rule",
		Severity:    "warning",
		Enabled:     true,
		Category:    "content",
	}
	
	validator.AddRule(rule)
	
	if len(validator.Rules) != 1 {
		t.Errorf("Expected 1 rule, got %d", len(validator.Rules))
	}
	
	if validator.Rules[0].ID != "test_rule" {
		t.Errorf("Expected rule ID 'test_rule', got '%s'", validator.Rules[0].ID)
	}
	
	if validator.Rules[0].Severity != "warning" {
		t.Errorf("Expected rule severity 'warning', got '%s'", validator.Rules[0].Severity)
	}
}

func TestValidateDocument(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 创建测试文档
	document := &wordprocessingml.Document{
		// 模拟文档结构
	}
	
	// 验证文档
	results, err := validator.ValidateDocument(document)
	if err != nil {
		t.Fatalf("Failed to validate document: %v", err)
	}
	
	if len(results) == 0 {
		t.Error("Expected validation results")
	}
}

func TestGetValidationResults(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 添加一些验证结果
	result1 := wordprocessingml.ValidationResult{
		RuleID:    "rule1",
		Severity:  "error",
		Message:   "Test error message",
		Location:  "document.xml",
		Line:      10,
		Column:    5,
	}
	
	result2 := wordprocessingml.ValidationResult{
		RuleID:    "rule2",
		Severity:  "warning",
		Message:   "Test warning message",
		Location:  "styles.xml",
		Line:      15,
		Column:    8,
	}
	
	validator.Results = append(validator.Results, result1, result2)
	
	results := validator.GetValidationResults()
	
	if len(results) != 2 {
		t.Errorf("Expected 2 validation results, got %d", len(results))
	}
	
	if results[0].RuleID != "rule1" {
		t.Errorf("Expected first result rule ID 'rule1', got '%s'", results[0].RuleID)
	}
	
	if results[1].Severity != "warning" {
		t.Errorf("Expected second result severity 'warning', got '%s'", results[1].Severity)
	}
}

func TestGetErrors(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 添加不同类型的验证结果
	result1 := wordprocessingml.ValidationResult{
		RuleID:   "rule1",
		Severity: "error",
		Message:  "Error message 1",
	}
	
	result2 := wordprocessingml.ValidationResult{
		RuleID:   "rule2",
		Severity: "warning",
		Message:  "Warning message",
	}
	
	result3 := wordprocessingml.ValidationResult{
		RuleID:   "rule3",
		Severity: "error",
		Message:  "Error message 2",
	}
	
	validator.Results = append(validator.Results, result1, result2, result3)
	
	errors := validator.GetErrors()
	
	if len(errors) != 2 {
		t.Errorf("Expected 2 errors, got %d", len(errors))
	}
	
	for _, err := range errors {
		if err.Severity != "error" {
			t.Errorf("Expected error severity 'error', got '%s'", err.Severity)
		}
	}
}

func TestGetWarnings(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 添加不同类型的验证结果
	result1 := wordprocessingml.ValidationResult{
		RuleID:   "rule1",
		Severity: "warning",
		Message:  "Warning message 1",
	}
	
	result2 := wordprocessingml.ValidationResult{
		RuleID:   "rule2",
		Severity: "error",
		Message:  "Error message",
	}
	
	result3 := wordprocessingml.ValidationResult{
		RuleID:   "rule3",
		Severity: "warning",
		Message:  "Warning message 2",
	}
	
	validator.Results = append(validator.Results, result1, result2, result3)
	
	warnings := validator.GetWarnings()
	
	if len(warnings) != 2 {
		t.Errorf("Expected 2 warnings, got %d", len(warnings))
	}
	
	for _, warning := range warnings {
		if warning.Severity != "warning" {
			t.Errorf("Expected warning severity 'warning', got '%s'", warning.Severity)
		}
	}
}

func TestGetInfo(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 添加不同类型的验证结果
	result1 := wordprocessingml.ValidationResult{
		RuleID:   "rule1",
		Severity: "info",
		Message:  "Info message 1",
	}
	
	result2 := wordprocessingml.ValidationResult{
		RuleID:   "rule2",
		Severity: "error",
		Message:  "Error message",
	}
	
	result3 := wordprocessingml.ValidationResult{
		RuleID:   "rule3",
		Severity: "info",
		Message:  "Info message 2",
	}
	
	validator.Results = append(validator.Results, result1, result2, result3)
	
	info := validator.GetInfo()
	
	if len(info) != 2 {
		t.Errorf("Expected 2 info messages, got %d", len(info))
	}
	
	for _, infoMsg := range info {
		if infoMsg.Severity != "info" {
			t.Errorf("Expected info severity 'info', got '%s'", infoMsg.Severity)
		}
	}
}

func TestHasErrors(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 测试没有错误的情况
	if validator.HasErrors() {
		t.Error("Expected no errors initially")
	}
	
	// 添加错误
	result := wordprocessingml.ValidationResult{
		RuleID:   "rule1",
		Severity: "error",
		Message:  "Test error",
	}
	validator.Results = append(validator.Results, result)
	
	// 测试有错误的情况
	if !validator.HasErrors() {
		t.Error("Expected to have errors")
	}
}

func TestHasWarnings(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 测试没有警告的情况
	if validator.HasWarnings() {
		t.Error("Expected no warnings initially")
	}
	
	// 添加警告
	result := wordprocessingml.ValidationResult{
		RuleID:   "rule1",
		Severity: "warning",
		Message:  "Test warning",
	}
	validator.Results = append(validator.Results, result)
	
	// 测试有警告的情况
	if !validator.HasWarnings() {
		t.Error("Expected to have warnings")
	}
}

func TestGetValidationSummary(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 添加不同类型的验证结果
	results := []wordprocessingml.ValidationResult{
		{RuleID: "rule1", Severity: "error", Message: "Error 1"},
		{RuleID: "rule2", Severity: "error", Message: "Error 2"},
		{RuleID: "rule3", Severity: "warning", Message: "Warning 1"},
		{RuleID: "rule4", Severity: "info", Message: "Info 1"},
	}
	
	validator.Results = results
	
	summary := validator.GetValidationSummary()
	
	if summary == "" {
		t.Error("Expected non-empty validation summary")
	}
	
	// 检查摘要是否包含预期的信息
	expectedInfo := []string{"2 errors", "1 warning", "1 info", "4 total"}
	for _, expected := range expectedInfo {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

func TestSetAutoFix(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 默认应该是false
	if validator.AutoFix {
		t.Error("Expected AutoFix to be false by default")
	}
	
	// 设置为true
	validator.SetAutoFix(true)
	
	if !validator.AutoFix {
		t.Error("Expected AutoFix to be true after setting")
	}
	
	// 设置为false
	validator.SetAutoFix(false)
	
	if validator.AutoFix {
		t.Error("Expected AutoFix to be false after setting")
	}
}

func TestValidateDocumentIntegrity(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 创建测试文档
	document := &wordprocessingml.Document{
		// 模拟文档结构
	}
	
	// 验证文档完整性
	results := validator.ValidateDocumentIntegrity(document)
	
	if len(results) == 0 {
		t.Error("Expected integrity validation results")
	}
	
	// 检查是否有完整性错误
	hasIntegrityErrors := false
	for _, result := range results {
		if result.Severity == "error" && contains(result.Message, "integrity") {
			hasIntegrityErrors = true
			break
		}
	}
	
	// 对于有效的文档，应该没有完整性错误
	if hasIntegrityErrors {
		t.Error("Expected document to pass integrity validation")
	}
}

func TestValidateContentStructure(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 创建测试文档
	document := &wordprocessingml.Document{
		// 模拟文档结构
	}
	
	// 验证内容结构
	results := validator.ValidateContentStructure(document)
	
	if len(results) == 0 {
		t.Error("Expected content structure validation results")
	}
	
	// 检查是否有结构错误
	hasStructureErrors := false
	for _, result := range results {
		if result.Severity == "error" && contains(result.Message, "structure") {
			hasStructureErrors = true
			break
		}
	}
	
	// 对于有效的文档，应该没有结构错误
	if hasStructureErrors {
		t.Error("Expected document to pass structure validation")
	}
}

func TestValidateFormatStandards(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 创建测试文档
	document := &wordprocessingml.Document{
		// 模拟文档结构
	}
	
	// 验证格式标准
	results := validator.ValidateFormatStandards(document)
	
	if len(results) == 0 {
		t.Error("Expected format standards validation results")
	}
	
	// 检查是否有格式错误
	hasFormatErrors := false
	for _, result := range results {
		if result.Severity == "error" && contains(result.Message, "format") {
			hasFormatErrors = true
			break
		}
	}
	
	// 对于有效的文档，应该没有格式错误
	if hasFormatErrors {
		t.Error("Expected document to pass format validation")
	}
}

func TestValidateTextQuality(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 创建测试文档
	document := &wordprocessingml.Document{
		// 模拟文档结构
	}
	
	// 验证文本质量
	results := validator.ValidateTextQuality(document)
	
	if len(results) == 0 {
		t.Error("Expected text quality validation results")
	}
	
	// 检查是否有文本质量错误
	hasQualityErrors := false
	for _, result := range results {
		if result.Severity == "error" && contains(result.Message, "quality") {
			hasQualityErrors = true
			break
		}
	}
	
	// 对于有效的文档，应该没有文本质量错误
	if hasQualityErrors {
		t.Error("Expected document to pass text quality validation")
	}
}

func TestValidateTableStructure(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 创建测试文档
	document := &wordprocessingml.Document{
		// 模拟文档结构
	}
	
	// 验证表格结构
	results := validator.ValidateTableStructure(document)
	
	if len(results) == 0 {
		t.Error("Expected table structure validation results")
	}
	
	// 检查是否有表格结构错误
	hasTableErrors := false
	for _, result := range results {
		if result.Severity == "error" && contains(result.Message, "table") {
			hasTableErrors = true
			break
		}
	}
	
	// 对于有效的文档，应该没有表格结构错误
	if hasTableErrors {
		t.Error("Expected document to pass table structure validation")
	}
}

func TestFixValidationIssues(t *testing.T) {
	validator := wordprocessingml.NewDocumentValidator()
	
	// 创建测试文档
	document := &wordprocessingml.Document{
		// 模拟文档结构
	}
	
	// 添加一些验证结果
	result := wordprocessingml.ValidationResult{
		RuleID:   "fixable_rule",
		Severity: "error",
		Message:  "Fixable error",
		Fixable:  true,
	}
	validator.Results = append(validator.Results, result)
	
	// 启用自动修复
	validator.SetAutoFix(true)
	
	// 修复验证问题
	fixedCount := validator.FixValidationIssues(document)
	
	if fixedCount == 0 {
		t.Error("Expected some issues to be fixed")
	}
}

// 辅助函数
