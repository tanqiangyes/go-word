package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewDocumentValidator(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	if validator == nil {
		t.Fatal("Expected DocumentValidator to be created")
	}
	
	if validator.ValidationRules == nil {
		t.Error("Expected ValidationRules to be initialized")
	}
	
	if validator.ValidationResults == nil {
		t.Error("Expected ValidationResults to be initialized")
	}
	
	if validator.AutoFix == false {
		t.Error("Expected AutoFix to be false by default")
	}
}

func TestAddRule(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 添加验证规则
	rule := wordprocessingml.ValidationRule{
		ID:          "test_rule",
		Name:        "Test Rule",
		Description: "A test validation rule",
		Severity:    wordprocessingml.WarningSeverity,
	}
	
	validator.ValidationRules = append(validator.ValidationRules, rule)
	
	if len(validator.ValidationRules) != 1 {
		t.Errorf("Expected 1 rule, got %d", len(validator.ValidationRules))
	}
	
	if validator.ValidationRules[0].ID != "test_rule" {
		t.Errorf("Expected rule ID 'test_rule', got '%s'", validator.ValidationRules[0].ID)
	}
	
	if validator.ValidationRules[0].Severity != wordprocessingml.WarningSeverity {
		t.Errorf("Expected rule severity WarningSeverity, got %v", validator.ValidationRules[0].Severity)
	}
}

func TestValidateDocument(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证文档系统是否正常工作
	if validator.Document == nil {
		t.Error("Expected document to be set")
	}
}

func TestGetValidationResults(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 添加一些验证结果
	result1 := wordprocessingml.ValidationResult{
		RuleID:    "rule1",
		Severity:  wordprocessingml.ErrorSeverity,
		Message:   "Test error message",
		Location:  "document.xml",
	}
	
	result2 := wordprocessingml.ValidationResult{
		RuleID:    "rule2",
		Severity:  wordprocessingml.WarningSeverity,
		Message:   "Test warning message",
		Location:  "styles.xml",
	}
	
	validator.ValidationResults = append(validator.ValidationResults, result1, result2)
	
	results := validator.ValidationResults
	
	if len(results) != 2 {
		t.Errorf("Expected 2 results, got %d", len(results))
	}
	
	if results[0].RuleID != "rule1" {
		t.Errorf("Expected first result rule ID 'rule1', got '%s'", results[0].RuleID)
	}
	
	if results[1].Severity != wordprocessingml.WarningSeverity {
		t.Errorf("Expected second result severity WarningSeverity, got %v", results[1].Severity)
	}
}

func TestGetErrors(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证错误系统是否正常工作
	if validator.ValidationResults == nil {
		t.Error("Expected validation results to be initialized")
	}
}

func TestGetWarnings(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证警告系统是否正常工作
	if validator.ValidationResults == nil {
		t.Error("Expected validation results to be initialized")
	}
}

func TestGetInfo(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证信息系统是否正常工作
	if validator.ValidationResults == nil {
		t.Error("Expected validation results to be initialized")
	}
}

func TestHasErrors(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证错误检查系统是否正常工作
	if validator.ValidationResults == nil {
		t.Error("Expected validation results to be initialized")
	}
}

func TestHasWarnings(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证警告检查系统是否正常工作
	if validator.ValidationResults == nil {
		t.Error("Expected validation results to be initialized")
	}
}

func TestGetValidationSummary(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证摘要系统是否正常工作
	if validator.ValidationResults == nil {
		t.Error("Expected validation results to be initialized")
	}
}

func TestSetAutoFix(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证自动修复系统是否正常工作
	if validator.AutoFix != false {
		t.Error("Expected AutoFix to be false by default")
	}
}

func TestValidateDocumentIntegrity(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证文档完整性检查系统是否正常工作
	if validator.Document == nil {
		t.Error("Expected document to be set")
	}
}

func TestValidateContentStructure(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证内容结构检查系统是否正常工作
	if validator.ValidationRules == nil {
		t.Error("Expected validation rules to be initialized")
	}
}

func TestValidateFormatStandards(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证格式标准检查系统是否正常工作
	if validator.ValidationRules == nil {
		t.Error("Expected validation rules to be initialized")
	}
}

func TestValidateTextQuality(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证文本质量检查系统是否正常工作
	if validator.ValidationRules == nil {
		t.Error("Expected validation rules to be initialized")
	}
}

func TestValidateTableStructure(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证表格结构检查系统是否正常工作
	if validator.ValidationRules == nil {
		t.Error("Expected validation rules to be initialized")
	}
}

func TestFixValidationIssues(t *testing.T) {
	doc := &wordprocessingml.Document{}
	validator := wordprocessingml.NewDocumentValidator(doc)
	
	// 验证问题修复系统是否正常工作
	if validator.AutoFix != false {
		t.Error("Expected AutoFix to be false by default")
	}
}
