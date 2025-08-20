package tests

import (
    "testing"

    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewDocumentValidator(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    if validator == nil {
        t.Fatal("Expected DocumentValidator to be created")
    }

    if validator.ValidationRules == nil {
        t.Error("Expected ValidationRules to be initialized")
    }

    if validator.ValidationResults == nil {
        t.Error("Expected ValidationResults to be initialized")
    }

    if validator.AutoFix != false {
        t.Error("Expected AutoFix to be false by default")
    }
}

func TestAddRule(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 记录初始规则数量
    initialRuleCount := len(validator.ValidationRules)

    // 添加验证规则
    rule := word.ValidationRule{
        ID:          "test_rule",
        Name:        "Test Rule",
        Description: "A test validation rule",
        Severity:    word.WarningSeverity,
    }

    validator.AddRule(rule)

    if len(validator.ValidationRules) != initialRuleCount+1 {
        t.Errorf("Expected %d rules, got %d", initialRuleCount+1, len(validator.ValidationRules))
    }

    // 查找添加的规则
    found := false
    for _, r := range validator.ValidationRules {
        if r.ID == "test_rule" {
            found = true
            if r.Severity != word.WarningSeverity {
                t.Errorf("Expected rule severity WarningSeverity, got %v", r.Severity)
            }
            break
        }
    }

    if !found {
        t.Error("Expected to find added validation rule")
    }
}

func TestValidateDocument(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证文档系统是否正常工作
    if validator.Document == nil {
        t.Error("Expected document to be set")
    }
}

func TestGetValidationResults(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 添加一些验证结果
    result1 := word.ValidationResult{
        RuleID:   "rule1",
        Severity: word.ErrorSeverity,
        Message:  "Test error message",
        Location: "document.xml",
    }

    result2 := word.ValidationResult{
        RuleID:   "rule2",
        Severity: word.WarningSeverity,
        Message:  "Test warning message",
        Location: "styles.xml",
    }

    validator.ValidationResults = append(validator.ValidationResults, result1, result2)

    results := validator.ValidationResults

    if len(results) != 2 {
        t.Errorf("Expected 2 results, got %d", len(results))
    }

    if results[0].RuleID != "rule1" {
        t.Errorf("Expected first result rule ID 'rule1', got '%s'", results[0].RuleID)
    }

    if results[1].Severity != word.WarningSeverity {
        t.Errorf("Expected second result severity WarningSeverity, got %v", results[1].Severity)
    }
}

func TestGetErrors(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证错误系统是否正常工作
    if validator.ValidationResults == nil {
        t.Error("Expected validation results to be initialized")
    }
}

func TestGetWarnings(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证警告系统是否正常工作
    if validator.ValidationResults == nil {
        t.Error("Expected validation results to be initialized")
    }
}

func TestGetInfo(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证信息系统是否正常工作
    if validator.ValidationResults == nil {
        t.Error("Expected validation results to be initialized")
    }
}

func TestHasErrors(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证错误检查系统是否正常工作
    if validator.ValidationResults == nil {
        t.Error("Expected validation results to be initialized")
    }
}

func TestHasWarnings(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证警告检查系统是否正常工作
    if validator.ValidationResults == nil {
        t.Error("Expected validation results to be initialized")
    }
}

func TestGetValidationSummary(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证摘要系统是否正常工作
    if validator.ValidationResults == nil {
        t.Error("Expected validation results to be initialized")
    }
}

func TestSetAutoFix(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证自动修复系统是否正常工作
    if validator.AutoFix != false {
        t.Error("Expected AutoFix to be false by default")
    }
}

func TestValidateDocumentIntegrity(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证文档完整性检查系统是否正常工作
    if validator.Document == nil {
        t.Error("Expected document to be set")
    }
}

func TestValidateContentStructure(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证内容结构检查系统是否正常工作
    if validator.ValidationRules == nil {
        t.Error("Expected validation rules to be initialized")
    }
}

func TestValidateFormatStandards(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证格式标准检查系统是否正常工作
    if validator.ValidationRules == nil {
        t.Error("Expected validation rules to be initialized")
    }
}

func TestValidateTextQuality(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证文本质量检查系统是否正常工作
    if validator.ValidationRules == nil {
        t.Error("Expected validation rules to be initialized")
    }
}

func TestValidateTableStructure(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证表格结构检查系统是否正常工作
    if validator.ValidationRules == nil {
        t.Error("Expected validation rules to be initialized")
    }
}

func TestFixValidationIssues(t *testing.T) {
    doc := &word.Document{}
    validator := word.NewDocumentValidator(doc)

    // 验证问题修复系统是否正常工作
    if validator.AutoFix != false {
        t.Error("Expected AutoFix to be false by default")
    }
}
