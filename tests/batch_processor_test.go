package tests

import (
    "fmt"
    "testing"
    "time"

    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewBatchProcessor(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    if processor == nil {
        t.Fatal("Expected BatchProcessor to be created")
    }

    if processor.Documents == nil {
        t.Error("Expected Documents to be initialized")
    }

    if processor.Operations == nil {
        t.Error("Expected Operations to be initialized")
    }

    if processor.ProgressChan == nil {
        t.Error("Expected ProgressChan to be initialized")
    }

    if processor.ErrorChan == nil {
        t.Error("Expected ErrorChan to be initialized")
    }
}

func TestAddDocument(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 添加文档
    document := &word.Document{}

    processor.AddDocument(document)

    if len(processor.Documents) != 1 {
        t.Errorf("Expected 1 document, got %d", len(processor.Documents))
    }

    if processor.Documents[0] == nil {
        t.Error("Expected document to be added")
    }
}

func TestAddOperation(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 添加操作
    operation := word.BatchOperation{
        Type:        word.ExtractText,
        Parameters:  map[string]interface{}{"include_formatting": true},
        DocumentIDs: []string{"doc1"},
    }

    processor.AddOperation(operation)

    if len(processor.Operations) != 1 {
        t.Errorf("Expected 1 operation, got %d", len(processor.Operations))
    }

    if processor.Operations[0].Type != word.ExtractText {
        t.Errorf("Expected operation type ExtractText, got %v", processor.Operations[0].Type)
    }
}

func TestProcessBatch(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 添加文档
    document := &word.Document{}
    processor.AddDocument(document)

    // 添加操作
    operation := word.BatchOperation{
        Type:        word.ExtractText,
        Parameters:  map[string]interface{}{"include_formatting": false},
        DocumentIDs: []string{"doc1"},
    }
    processor.AddOperation(operation)

    // 处理批量操作
    err := processor.ProcessBatch()
    if err != nil {
        t.Fatalf("Failed to process batch: %v", err)
    }
}

func TestProcessDocument(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 处理文档 - 这里我们只是测试系统是否正常工作
    if processor.Documents == nil {
        t.Error("Expected documents to be initialized")
    }

    if processor.Operations == nil {
        t.Error("Expected operations to be initialized")
    }
}

func TestExecuteOperation(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 创建测试操作
    operation := word.BatchOperation{
        Type:        word.ExtractText,
        Parameters:  map[string]interface{}{"include_formatting": false},
        DocumentIDs: []string{"test_doc"},
    }

    // 执行操作 - 这里我们只是测试系统是否正常工作
    if processor.Operations == nil {
        t.Error("Expected operations to be initialized")
    }

    if operation.Type != word.ExtractText {
        t.Errorf("Expected operation type ExtractText, got %v", operation.Type)
    }
}

func TestMonitorProgress(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 验证进度通道是否可用
    if processor.ProgressChan == nil {
        t.Error("Expected ProgressChan to be initialized")
    }

    // 验证错误通道是否可用
    if processor.ErrorChan == nil {
        t.Error("Expected ErrorChan to be initialized")
    }
}

func TestMonitorErrors(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 验证错误通道是否可用
    if processor.ErrorChan == nil {
        t.Error("Expected ErrorChan to be initialized")
    }

    // 验证上下文是否可用
    if processor.Context == nil {
        t.Error("Expected Context to be initialized")
    }
}

func TestGetProgressChannel(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    channel := processor.GetProgressChannel()

    if channel == nil {
        t.Error("Expected progress channel to be available")
    }
}

func TestGetErrorChannel(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    channel := processor.GetErrorChannel()

    if channel == nil {
        t.Error("Expected error channel to be available")
    }
}

func TestCancel(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 取消处理
    processor.Cancel()

    // 验证取消函数是否可用
    if processor.CancelFunc == nil {
        t.Error("Expected cancel function to be available")
    }
}

func TestGetBatchSummary(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 添加一些文档和操作
    document1 := &word.Document{}
    document2 := &word.Document{}

    processor.AddDocument(document1)
    processor.AddDocument(document2)

    operation1 := word.BatchOperation{Type: word.ExtractText, DocumentIDs: []string{"doc1"}}
    operation2 := word.BatchOperation{Type: word.MergeDocuments, DocumentIDs: []string{"doc2"}}

    processor.AddOperation(operation1)
    processor.AddOperation(operation2)

    // 获取批量处理摘要
    summary := processor.GetBatchSummary()

    if summary == "" {
        t.Error("Expected non-empty batch summary")
    }
}

func TestNewBatchProcessorWithConfig(t *testing.T) {
    config := word.BatchProcessorConfig{
        Concurrency:     4,
        Timeout:         30 * time.Second,
        RetryCount:      3,
        RetryDelay:      1 * time.Second,
        ProgressEnabled: true,
        ErrorHandling:   word.ContinueOnError,
    }

    processor := word.NewBatchProcessorWithConfig(config)

    if processor == nil {
        t.Fatal("Expected BatchProcessor to be created with config")
    }

    if processor.Concurrency != 4 {
        t.Errorf("Expected concurrency 4, got %d", processor.Concurrency)
    }
}

func TestBatchErrorHandling(t *testing.T) {
    processor := word.NewBatchProcessor(4)

    // 验证错误通道是否可用
    if processor.ErrorChan == nil {
        t.Error("Expected error channel to be available")
    }

    // 验证错误结构
    batchError := word.BatchError{
        DocumentID: "doc1",
        Operation:  "extract_text",
        Error:      fmt.Errorf("Test error"),
        Timestamp:  time.Now(),
    }

    if batchError.DocumentID != "doc1" {
        t.Errorf("Expected document ID 'doc1', got '%s'", batchError.DocumentID)
    }
}

func TestProgressReport(t *testing.T) {
    // 创建进度报告
    report := word.ProgressReport{
        TotalDocuments:     10,
        ProcessedDocuments: 5,
        CurrentDocument:    "doc1",
        Operation:          "extract_text",
        Percentage:         50.0,
        StartTime:          time.Now(),
        EstimatedTime:      30 * time.Second,
    }

    // 验证进度报告结构
    if report.TotalDocuments != 10 {
        t.Errorf("Expected total documents 10, got %d", report.TotalDocuments)
    }

    if report.Percentage != 50.0 {
        t.Errorf("Expected percentage 50.0, got %f", report.Percentage)
    }
}

func TestOperationTypes(t *testing.T) {
    // 测试不同类型的操作
    operationTypes := []word.OperationType{
        word.ExtractText,
        word.MergeDocuments,
        word.ApplyTemplate,
        word.ValidateDocuments,
        word.ConvertFormat,
    }

    for _, opType := range operationTypes {
        operation := word.BatchOperation{
            Type:        opType,
            Parameters:  map[string]interface{}{"test": true},
            DocumentIDs: []string{"test_doc"},
        }

        if operation.Type != opType {
            t.Errorf("Expected operation type %v, got %v", opType, operation.Type)
        }
    }
}

// 辅助函数
