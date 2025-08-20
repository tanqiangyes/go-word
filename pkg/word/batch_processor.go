// Package word provides word document processing functionality
package word

import (
	"context"
	"fmt"
	"sync"
	"time"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// BatchProcessor represents a batch document processor
type BatchProcessor struct {
	Documents    []*Document
	Operations   []BatchOperation
	Concurrency  int
	ProgressChan chan ProgressReport
	ErrorChan    chan BatchError
	Context      context.Context
	CancelFunc   context.CancelFunc
}

// BatchOperation represents a single batch operation
type BatchOperation struct {
	Type        OperationType
	Parameters  map[string]interface{}
	DocumentIDs []string
}

// OperationType defines the type of batch operation
type OperationType int

const (
	// ExtractText extracts text from documents
	ExtractText OperationType = iota
	// ExtractTables extracts tables from documents
	ExtractTables
	// MergeDocuments merges multiple documents
	MergeDocuments
	// ApplyTemplate applies template to documents
	ApplyTemplate
	// ValidateDocuments validates documents
	ValidateDocuments
	// ConvertFormat converts document format
	ConvertFormat
)

// ProgressReport represents a progress report
type ProgressReport struct {
	TotalDocuments     int
	ProcessedDocuments int
	CurrentDocument    string
	Operation          string
	Percentage         float64
	StartTime          time.Time
	EstimatedTime      time.Duration
}

// BatchError represents a batch processing error
type BatchError struct {
	DocumentID string
	Operation  string
	Error      error
	Timestamp  time.Time
}

// NewBatchProcessor creates a new batch processor
func NewBatchProcessor(concurrency int) *BatchProcessor {
	ctx, cancel := context.WithCancel(context.Background())

	return &BatchProcessor{
		Documents:    make([]*Document, 0),
		Operations:   make([]BatchOperation, 0),
		Concurrency:  concurrency,
		ProgressChan: make(chan ProgressReport, 100),
		ErrorChan:    make(chan BatchError, 100),
		Context:      ctx,
		CancelFunc:   cancel,
	}
}

// AddDocument adds a document to the batch processor
func (bp *BatchProcessor) AddDocument(doc *Document) {
	bp.Documents = append(bp.Documents, doc)
}

// AddOperation adds an operation to the batch processor
func (bp *BatchProcessor) AddOperation(operation BatchOperation) {
	bp.Operations = append(bp.Operations, operation)
}

// ProcessBatch processes all documents with all operations
func (bp *BatchProcessor) ProcessBatch() error {
	if len(bp.Documents) == 0 {
		return fmt.Errorf("no documents to process")
	}

	if len(bp.Operations) == 0 {
		return fmt.Errorf("no operations to perform")
	}

	// 启动进度监控
	go bp.monitorProgress()

	// 创建工作池
	workerPool := make(chan struct{}, bp.Concurrency)
	var wg sync.WaitGroup

	// 为每个文档启动一个工作协程
	for i, doc := range bp.Documents {
		select {
		case <-bp.Context.Done():
			return fmt.Errorf("batch processing cancelled")
		case workerPool <- struct{}{}:
			wg.Add(1)
			go func(docIndex int, document *Document) {
				defer wg.Done()
				defer func() { <-workerPool }()

				bp.processDocument(docIndex, document)
			}(i, doc)
		}
	}

	// 等待所有工作完成
	wg.Wait()

	// 通知监控协程退出，避免在关闭通道期间仍尝试发送
	// 这里不立即关闭通道以避免竞态条件导致的 panic: send on closed channel
	bp.CancelFunc()

	return nil
}

// processDocument processes a single document
func (bp *BatchProcessor) processDocument(docIndex int, doc *Document) {
	docID := fmt.Sprintf("doc_%d", docIndex)

	for _, operation := range bp.Operations {
		select {
		case <-bp.Context.Done():
			return
		default:
			if err := bp.executeOperation(docID, doc, operation); err != nil {
				bp.ErrorChan <- BatchError{
					DocumentID: docID,
					Operation:  operation.Type.String(),
					Error:      err,
					Timestamp:  time.Now(),
				}
			}
		}
	}
}

// executeOperation executes a single operation on a document
func (bp *BatchProcessor) executeOperation(docID string, doc *Document, operation BatchOperation) error {
	switch operation.Type {
	case ExtractText:
		return bp.extractTextFromDocument(docID, doc)
	case ExtractTables:
		return bp.extractTablesFromDocument(docID, doc)
	case MergeDocuments:
		return bp.mergeDocuments(docID, doc, operation)
	case ApplyTemplate:
		return bp.applyTemplateToDocument(docID, doc, operation)
	case ValidateDocuments:
		return bp.validateDocument(docID, doc)
	case ConvertFormat:
		return bp.convertDocumentFormat(docID, doc, operation)
	default:
		return fmt.Errorf("unsupported operation type: %v", operation.Type)
	}
}

// extractTextFromDocument extracts text from a document
func (bp *BatchProcessor) extractTextFromDocument(docID string, doc *Document) error {
	_, err := doc.GetText()
	if err != nil {
		return fmt.Errorf("failed to extract text from %s: %w", docID, err)
	}

	// 这里可以添加文本处理逻辑
	// 例如：保存到文件、发送到API等

	return nil
}

// extractTablesFromDocument extracts tables from a document
func (bp *BatchProcessor) extractTablesFromDocument(docID string, doc *Document) error {
	_, err := doc.GetTables()
	if err != nil {
		return fmt.Errorf("failed to extract tables from %s: %w", docID, err)
	}

	// 这里可以添加表格处理逻辑
	// 例如：转换为CSV、保存到数据库等

	return nil
}

// mergeDocuments merges documents
func (bp *BatchProcessor) mergeDocuments(docID string, doc *Document, operation BatchOperation) error {
	// 实现文档合并逻辑
	merge := NewDocumentMerge(doc)

	// 添加源文档
	for range operation.DocumentIDs {
		// 这里需要根据ID查找源文档
		// 暂时跳过实现
	}

	return merge.MergeDocuments()
}

// applyTemplateToDocument applies template to a document
func (bp *BatchProcessor) applyTemplateToDocument(docID string, doc *Document, operation BatchOperation) error {
	template := NewTemplate(doc)

	// 从参数中获取变量
	if variables, ok := operation.Parameters["variables"].(map[string]interface{}); ok {
		for key, value := range variables {
			template.AddVariable(key, value)
		}
	}

	return template.ProcessTemplate()
}

// validateDocument validates a document
func (bp *BatchProcessor) validateDocument(docID string, doc *Document) error {
	// 实现文档验证逻辑
	if doc.mainPart == nil {
		return fmt.Errorf("document %s has no main part", docID)
	}

	if doc.mainPart.Content == nil {
		return fmt.Errorf("document %s has no content", docID)
	}

	// 检查段落数量
	if len(doc.mainPart.Content.Paragraphs) == 0 {
		return fmt.Errorf("document %s has no paragraphs", docID)
	}

	return nil
}

// convertDocumentFormat converts document format
func (bp *BatchProcessor) convertDocumentFormat(docID string, doc *Document, operation BatchOperation) error {
	// 实现格式转换逻辑
	targetFormat, ok := operation.Parameters["target_format"].(string)
	if !ok {
		return fmt.Errorf("target format not specified")
	}

	switch targetFormat {
	case "txt":
		return bp.convertToText(docID, doc)
	case "html":
		return bp.convertToHTML(docID, doc)
	case "pdf":
		return bp.convertToPDF(docID, doc)
	default:
		return fmt.Errorf("unsupported target format: %s", targetFormat)
	}
}

// convertToText converts document to text format
func (bp *BatchProcessor) convertToText(docID string, doc *Document) error {
	_, err := doc.GetText()
	if err != nil {
		return fmt.Errorf("failed to get text from %s: %w", docID, err)
	}

	// 这里可以保存为文本文件
	// 暂时只记录日志

	return nil
}

// convertToHTML converts document to HTML format
func (bp *BatchProcessor) convertToHTML(docID string, doc *Document) error {
	// 获取文档内容
	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		return fmt.Errorf("获取文档段落失败: %w", err)
	}

	tables, err := doc.GetTables()
	if err != nil {
		// 如果获取表格失败，继续处理但不包含表格
		tables = []types.Table{}
	}

	// 构建HTML内容
	htmlContent := bp.buildHTMLContent(paragraphs, tables)

	// 这里可以保存HTML文件或进行其他处理
	_ = htmlContent

	return nil
}

// convertToPDF converts document to PDF format
func (bp *BatchProcessor) convertToPDF(docID string, doc *Document) error {
	// 先转换为HTML，然后基于HTML生成PDF
	if err := bp.convertToHTML(docID, doc); err != nil {
		return fmt.Errorf("转换为HTML失败: %w", err)
	}

	// 获取文档文本内容用于PDF生成
	text, err := doc.GetText()
	if err != nil {
		return fmt.Errorf("获取文档文本失败: %w", err)
	}

	// 生成简单的PDF内容（这里使用基本的PDF格式）
	pdfContent := bp.generateBasicPDF(text)

	// 这里可以保存PDF文件或进行其他处理
	_ = pdfContent

	return nil
}

// buildHTMLContent 构建HTML内容
func (bp *BatchProcessor) buildHTMLContent(paragraphs []types.Paragraph, tables []types.Table) string {
	html := `<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
	<title>Document</title>
	<style>
		body { font-family: Arial, sans-serif; margin: 20px; }
		p { margin: 10px 0; }
		table { border-collapse: collapse; width: 100%; margin: 10px 0; }
		th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
		th { background-color: #f2f2f2; }
	</style>
</head>
<body>
`

	// 添加段落
	for _, paragraph := range paragraphs {
		html += fmt.Sprintf("\t<p>%s</p>\n", paragraph.Text)
	}

	// 添加表格
	for _, table := range tables {
		html += "\t<table>\n"
		for i, row := range table.Rows {
			if i == 0 {
				// 第一行作为表头
				html += "\t\t<tr>\n"
				for _, cell := range row.Cells {
					html += fmt.Sprintf("\t\t\t<th>%s</th>\n", cell.Text)
				}
				html += "\t\t</tr>\n"
			} else {
				// 其他行作为数据行
				html += "\t\t<tr>\n"
				for _, cell := range row.Cells {
					html += fmt.Sprintf("\t\t\t<td>%s</td>\n", cell.Text)
				}
				html += "\t\t</tr>\n"
			}
		}
		html += "\t</table>\n"
	}

	html += "</body>\n</html>"
	return html
}

// generateBasicPDF 生成基本的PDF内容
func (bp *BatchProcessor) generateBasicPDF(text string) string {
	// 这是一个非常简化的PDF格式，实际应用中应该使用专业的PDF库
	pdfHeader := "%PDF-1.4\n"
	pdfCatalog := "1 0 obj\n<<\n/Type /Catalog\n/Pages 2 0 R\n>>\nendobj\n"
	pdfPages := "2 0 obj\n<<\n/Type /Pages\n/Kids [3 0 R]\n/Count 1\n>>\nendobj\n"
	pdfPage := "3 0 obj\n<<\n/Type /Page\n/Parent 2 0 R\n/MediaBox [0 0 612 792]\n/Contents 4 0 R\n>>\nendobj\n"
	pdfContent := fmt.Sprintf("4 0 obj\n<<\n/Length %d\n>>\nstream\nBT\n/F1 12 Tf\n72 720 Td\n(%s) Tj\nET\nendstream\nendobj\n", len(text)+30, text)
	pdfXref := "xref\n0 5\n0000000000 65535 f \n0000000009 00000 n \n0000000074 00000 n \n0000000120 00000 n \n0000000179 00000 n \n"
	pdfTrailer := fmt.Sprintf("trailer\n<<\n/Size 5\n/Root 1 0 R\n>>\nstartxref\n%d\n%%%%EOF", len(pdfHeader+pdfCatalog+pdfPages+pdfPage+pdfContent))

	return pdfHeader + pdfCatalog + pdfPages + pdfPage + pdfContent + pdfXref + pdfTrailer
}

// monitorProgress monitors the progress of batch processing
func (bp *BatchProcessor) monitorProgress() {
	totalDocs := len(bp.Documents)
	processedDocs := 0
	startTime := time.Now()

	ticker := time.NewTicker(1 * time.Second)
	defer ticker.Stop()

	for {
		select {
		case <-bp.Context.Done():
			return
		case <-ticker.C:
			processedDocs++
			percentage := float64(processedDocs) / float64(totalDocs) * 100

			elapsed := time.Since(startTime)
			var estimatedTime time.Duration
			if processedDocs > 0 {
				estimatedTime = time.Duration(float64(elapsed) * float64(totalDocs) / float64(processedDocs))
			}

			report := ProgressReport{
				TotalDocuments:     totalDocs,
				ProcessedDocuments: processedDocs,
				CurrentDocument:    fmt.Sprintf("doc_%d", processedDocs-1),
				Operation:          "batch_processing",
				Percentage:         percentage,
				StartTime:          startTime,
				EstimatedTime:      estimatedTime,
			}

			select {
			case bp.ProgressChan <- report:
			default:
				// 通道已满，跳过
			case <-bp.Context.Done():
				return
			}

			if processedDocs >= totalDocs {
				return
			}
		}
	}
}

// GetProgressChannel returns the progress channel
func (bp *BatchProcessor) GetProgressChannel() <-chan ProgressReport {
	return bp.ProgressChan
}

// GetErrorChannel returns the error channel
func (bp *BatchProcessor) GetErrorChannel() <-chan BatchError {
	return bp.ErrorChan
}

// Cancel cancels the batch processing
func (bp *BatchProcessor) Cancel() {
	bp.CancelFunc()
}

// GetBatchSummary returns a summary of the batch operation
func (bp *BatchProcessor) GetBatchSummary() string {
	var summary string
	summary += fmt.Sprintf("文档数量: %d\n", len(bp.Documents))
	summary += fmt.Sprintf("操作数量: %d\n", len(bp.Operations))
	summary += fmt.Sprintf("并发数: %d\n", bp.Concurrency)

	summary += "\n操作列表:\n"
	for i, operation := range bp.Operations {
		summary += fmt.Sprintf("%d. %s\n", i+1, operation.Type.String())
	}

	return summary
}

// String returns the string representation of OperationType
func (ot OperationType) String() string {
	switch ot {
	case ExtractText:
		return "ExtractText"
	case ExtractTables:
		return "ExtractTables"
	case MergeDocuments:
		return "MergeDocuments"
	case ApplyTemplate:
		return "ApplyTemplate"
	case ValidateDocuments:
		return "ValidateDocuments"
	case ConvertFormat:
		return "ConvertFormat"
	default:
		return "Unknown"
	}
}

// BatchProcessorConfig represents configuration for batch processor
type BatchProcessorConfig struct {
	Concurrency     int
	Timeout         time.Duration
	RetryCount      int
	RetryDelay      time.Duration
	ProgressEnabled bool
	ErrorHandling   ErrorHandlingMode
}

// ErrorHandlingMode defines how errors should be handled
type ErrorHandlingMode int

const (
	// StopOnError stops processing on first error
	StopOnError ErrorHandlingMode = iota
	// ContinueOnError continues processing despite errors
	ContinueOnError
	// RetryOnError retries failed operations
	RetryOnError
)

// NewBatchProcessorWithConfig creates a new batch processor with configuration
func NewBatchProcessorWithConfig(config BatchProcessorConfig) *BatchProcessor {
	concurrency := config.Concurrency
	if concurrency <= 0 {
		concurrency = 1
	}

	processor := NewBatchProcessor(concurrency)

	// 应用配置
	if config.Timeout > 0 {
		ctx, cancel := context.WithTimeout(context.Background(), config.Timeout)
		processor.Context = ctx
		processor.CancelFunc = cancel
	}

	return processor
}
