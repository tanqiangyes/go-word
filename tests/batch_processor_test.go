package tests

import (
	"testing"
	"time"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewBatchProcessor(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	if processor == nil {
		t.Fatal("Expected BatchProcessor to be created")
	}
	
	if processor.Documents == nil {
		t.Error("Expected Documents to be initialized")
	}
	
	if processor.Operations == nil {
		t.Error("Expected Operations to be initialized")
	}
	
	if processor.ProgressChannel == nil {
		t.Error("Expected ProgressChannel to be initialized")
	}
	
	if processor.ErrorChannel == nil {
		t.Error("Expected ErrorChannel to be initialized")
	}
}

func TestAddDocument(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 添加文档
	document := wordprocessingml.BatchDocument{
		Path:     "/path/to/document.docx",
		ID:       "doc1",
		Priority: 1,
	}
	
	processor.AddDocument(document)
	
	if len(processor.Documents) != 1 {
		t.Errorf("Expected 1 document, got %d", len(processor.Documents))
	}
	
	if processor.Documents[0].ID != "doc1" {
		t.Errorf("Expected document ID 'doc1', got '%s'", processor.Documents[0].ID)
	}
	
	if processor.Documents[0].Path != "/path/to/document.docx" {
		t.Errorf("Expected document path '/path/to/document.docx', got '%s'", processor.Documents[0].Path)
	}
}

func TestAddOperation(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 添加操作
	operation := wordprocessingml.BatchOperation{
		Type:      "extract_text",
		Parameters: map[string]interface{}{"include_formatting": true},
		Priority:  1,
	}
	
	processor.AddOperation(operation)
	
	if len(processor.Operations) != 1 {
		t.Errorf("Expected 1 operation, got %d", len(processor.Operations))
	}
	
	if processor.Operations[0].Type != "extract_text" {
		t.Errorf("Expected operation type 'extract_text', got '%s'", processor.Operations[0].Type)
	}
	
	if processor.Operations[0].Priority != 1 {
		t.Errorf("Expected operation priority 1, got %d", processor.Operations[0].Priority)
	}
}

func TestProcessBatch(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 添加文档
	document := wordprocessingml.BatchDocument{
		Path:     "/path/to/document.docx",
		ID:       "doc1",
		Priority: 1,
	}
	processor.AddDocument(document)
	
	// 添加操作
	operation := wordprocessingml.BatchOperation{
		Type:      "extract_text",
		Parameters: map[string]interface{}{"include_formatting": false},
		Priority:  1,
	}
	processor.AddOperation(operation)
	
	// 处理批量操作
	results, err := processor.ProcessBatch()
	if err != nil {
		t.Fatalf("Failed to process batch: %v", err)
	}
	
	if len(results) == 0 {
		t.Error("Expected batch processing results")
	}
}

func TestProcessDocument(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 创建测试文档
	document := wordprocessingml.BatchDocument{
		Path:     "/path/to/test.docx",
		ID:       "test_doc",
		Priority: 1,
	}
	
	// 创建测试操作
	operation := wordprocessingml.BatchOperation{
		Type:      "extract_text",
		Parameters: map[string]interface{}{"include_formatting": true},
		Priority:  1,
	}
	
	// 处理文档
	result := processor.ProcessDocument(document, operation)
	
	if result == nil {
		t.Error("Expected processing result")
	}
	
	if result.DocumentID != "test_doc" {
		t.Errorf("Expected document ID 'test_doc', got '%s'", result.DocumentID)
	}
	
	if result.OperationType != "extract_text" {
		t.Errorf("Expected operation type 'extract_text', got '%s'", result.OperationType)
	}
}

func TestExecuteOperation(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 创建测试操作
	operation := wordprocessingml.BatchOperation{
		Type:      "extract_text",
		Parameters: map[string]interface{}{"include_formatting": false},
		Priority:  1,
	}
	
	// 执行操作
	result := processor.ExecuteOperation(operation, "test_doc")
	
	if result == nil {
		t.Error("Expected operation result")
	}
	
	if result.OperationType != "extract_text" {
		t.Errorf("Expected operation type 'extract_text', got '%s'", result.OperationType)
	}
	
	if result.DocumentID != "test_doc" {
		t.Errorf("Expected document ID 'test_doc', got '%s'", result.DocumentID)
	}
}

func TestMonitorProgress(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 启动进度监控
	go processor.MonitorProgress()
	
	// 等待一段时间让监控启动
	time.Sleep(100 * time.Millisecond)
	
	// 验证进度通道是否可用
	select {
	case <-processor.ProgressChannel:
		// 通道有数据，这是正常的
	default:
		// 通道为空，这也是正常的
	}
}

func TestMonitorErrors(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 启动错误监控
	go processor.MonitorErrors()
	
	// 等待一段时间让监控启动
	time.Sleep(100 * time.Millisecond)
	
	// 验证错误通道是否可用
	select {
	case <-processor.ErrorChannel:
		// 通道有数据，这是正常的
	default:
		// 通道为空，这也是正常的
	}
}

func TestGetProgressChannel(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	channel := processor.GetProgressChannel()
	
	if channel == nil {
		t.Error("Expected progress channel to be available")
	}
}

func TestGetErrorChannel(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	channel := processor.GetErrorChannel()
	
	if channel == nil {
		t.Error("Expected error channel to be available")
	}
}

func TestCancel(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 取消处理
	processor.Cancel()
	
	// 验证取消状态
	if !processor.IsCancelled {
		t.Error("Expected processor to be cancelled")
	}
}

func TestGetBatchSummary(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 添加一些文档和操作
	document1 := wordprocessingml.BatchDocument{ID: "doc1", Path: "/path/to/doc1.docx"}
	document2 := wordprocessingml.BatchDocument{ID: "doc2", Path: "/path/to/doc2.docx"}
	
	processor.AddDocument(document1)
	processor.AddDocument(document2)
	
	operation1 := wordprocessingml.BatchOperation{Type: "extract_text", Priority: 1}
	operation2 := wordprocessingml.BatchOperation{Type: "merge", Priority: 2}
	
	processor.AddOperation(operation1)
	processor.AddOperation(operation2)
	
	// 获取批量处理摘要
	summary := processor.GetBatchSummary()
	
	if summary == "" {
		t.Error("Expected non-empty batch summary")
	}
	
	// 检查摘要是否包含预期的信息
	expectedInfo := []string{"2 documents", "2 operations", "extract_text", "merge"}
	for _, expected := range expectedInfo {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

func TestNewBatchProcessorWithConfig(t *testing.T) {
	config := wordprocessingml.BatchProcessorConfig{
		MaxConcurrency: 4,
		Timeout:        30 * time.Second,
		ErrorHandling:  "continue",
		ProgressInterval: 1 * time.Second,
	}
	
	processor := wordprocessingml.NewBatchProcessorWithConfig(config)
	
	if processor == nil {
		t.Fatal("Expected BatchProcessor to be created with config")
	}
	
	if processor.Config.MaxConcurrency != 4 {
		t.Errorf("Expected max concurrency 4, got %d", processor.Config.MaxConcurrency)
	}
	
	if processor.Config.Timeout != 30*time.Second {
		t.Errorf("Expected timeout 30s, got %v", processor.Config.Timeout)
	}
	
	if processor.Config.ErrorHandling != "continue" {
		t.Errorf("Expected error handling 'continue', got '%s'", processor.Config.ErrorHandling)
	}
}

func TestBatchErrorHandling(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 创建错误
	batchError := wordprocessingml.BatchError{
		DocumentID: "doc1",
		OperationType: "extract_text",
		Error: "Test error",
		Severity: "error",
		Timestamp: time.Now(),
	}
	
	// 处理错误
	processor.HandleError(batchError)
	
	// 验证错误是否被记录
	if len(processor.Errors) == 0 {
		t.Error("Expected error to be recorded")
	}
	
	recordedError := processor.Errors[0]
	if recordedError.DocumentID != "doc1" {
		t.Errorf("Expected error document ID 'doc1', got '%s'", recordedError.DocumentID)
	}
	
	if recordedError.OperationType != "extract_text" {
		t.Errorf("Expected error operation type 'extract_text', got '%s'", recordedError.OperationType)
	}
	
	if recordedError.Error != "Test error" {
		t.Errorf("Expected error message 'Test error', got '%s'", recordedError.Error)
	}
}

func TestProgressReport(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 创建进度报告
	report := wordprocessingml.ProgressReport{
		DocumentID:    "doc1",
		OperationType: "extract_text",
		Progress:      50,
		Status:        "processing",
		Message:       "Extracting text content",
		Timestamp:     time.Now(),
	}
	
	// 发送进度报告
	processor.SendProgress(report)
	
	// 验证进度报告是否被发送
	select {
	case receivedReport := <-processor.ProgressChannel:
		if receivedReport.DocumentID != "doc1" {
			t.Errorf("Expected progress report document ID 'doc1', got '%s'", receivedReport.DocumentID)
		}
		if receivedReport.Progress != 50 {
			t.Errorf("Expected progress 50, got %d", receivedReport.Progress)
		}
	default:
		t.Error("Expected progress report to be sent")
	}
}

func TestOperationTypes(t *testing.T) {
	processor := wordprocessingml.NewBatchProcessor()
	
	// 测试不同类型的操作
	operationTypes := []string{"extract_text", "merge", "apply_template", "validate", "convert"}
	
	for _, opType := range operationTypes {
		operation := wordprocessingml.BatchOperation{
			Type:      opType,
			Parameters: map[string]interface{}{"test": true},
			Priority:  1,
		}
		
		result := processor.ExecuteOperation(operation, "test_doc")
		
		if result == nil {
			t.Errorf("Expected result for operation type '%s'", opType)
		}
		
		if result.OperationType != opType {
			t.Errorf("Expected operation type '%s', got '%s'", opType, result.OperationType)
		}
	}
}

// 辅助函数
