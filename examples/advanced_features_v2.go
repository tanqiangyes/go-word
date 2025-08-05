package examples

import (
	"fmt"
	"log"
	"time"

	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

func DemoAdvancedFeaturesV2() {
	fmt.Println("=== Go Word 高级功能示例 V2 ===\n")

	// 示例1：批量文档处理
	demoBatchProcessing()

	// 示例2：文档验证
	demoDocumentValidation()

	// 示例3：性能优化
	demoPerformanceOptimization()

	fmt.Println("所有高级功能示例完成！")
}

// demoBatchProcessing 演示批量文档处理功能
func demoBatchProcessing() {
	fmt.Println("1. 批量文档处理示例")
	fmt.Println("-------------------")

	// 创建多个测试文档
	documents := createTestDocuments(5)

	// 创建批量处理器
	processor := wordprocessingml.NewBatchProcessor(3) // 3个并发

	// 添加文档
	for _, doc := range documents {
		processor.AddDocument(doc)
	}

	// 添加操作
	processor.AddOperation(wordprocessingml.BatchOperation{
		Type: wordprocessingml.ExtractText,
		Parameters: map[string]interface{}{
			"save_to_file": true,
		},
	})

	processor.AddOperation(wordprocessingml.BatchOperation{
		Type: wordprocessingml.ValidateDocuments,
		Parameters: map[string]interface{}{
			"auto_fix": true,
		},
	})

	processor.AddOperation(wordprocessingml.BatchOperation{
		Type: wordprocessingml.ConvertFormat,
		Parameters: map[string]interface{}{
			"target_format": "txt",
		},
	})

	// 启动进度监控
	go monitorProgress(processor)

	// 启动错误监控
	go monitorErrors(processor)

	// 执行批量处理
	fmt.Println("开始批量处理...")
	fmt.Println(processor.GetBatchSummary())

	if err := processor.ProcessBatch(); err != nil {
		log.Printf("批量处理失败: %v", err)
		return
	}

	fmt.Println("批量处理完成！")
	fmt.Println()
}

// demoDocumentValidation 演示文档验证功能
func demoDocumentValidation() {
	fmt.Println("2. 文档验证示例")
	fmt.Println("----------------")

	// 创建一个有问题的文档
	doc := createProblematicDocument()

	// 创建文档验证器
	validator := wordprocessingml.NewDocumentValidator(doc)

	// 启用自动修复
	validator.SetAutoFix(true)

	// 执行验证
	fmt.Println("开始文档验证...")
	if err := validator.ValidateDocument(); err != nil {
		log.Printf("文档验证失败: %v", err)
		return
	}

	// 显示验证结果
	fmt.Println("验证结果:")
	fmt.Println(validator.GetValidationSummary())

	// 显示详细结果
	results := validator.GetValidationResults()
	for _, result := range results {
		status := "通过"
		if result.Error != nil {
			status = "失败"
			if result.Fixed {
				status = "已修复"
			}
		}
		fmt.Printf("- %s (%s): %s\n", result.RuleID, result.Severity.String(), status)
	}

	fmt.Println()
}

// demoPerformanceOptimization 演示性能优化功能
func demoPerformanceOptimization() {
	fmt.Println("3. 性能优化示例")
	fmt.Println("----------------")

	// 创建大文档
	doc := createLargeDocument()

	// 测试解析性能
	fmt.Println("测试文档解析性能...")
	startTime := time.Now()

	// 解析文档
	text, err := doc.GetText()
	if err != nil {
		log.Printf("获取文本失败: %v", err)
		return
	}

	parseTime := time.Since(startTime)
	fmt.Printf("解析时间: %v\n", parseTime)
	fmt.Printf("文档大小: %d 字符\n", len(text))

	// 测试内存使用
	fmt.Println("测试内存使用...")
	// 这里可以添加内存使用统计
	// 在实际应用中可以使用 runtime.ReadMemStats

	// 测试并发处理
	fmt.Println("测试并发处理...")
	startTime = time.Now()

	// 模拟并发处理
	results := make(chan string, 10)
	for i := 0; i < 5; i++ {
		go func(id int) {
			// 模拟处理
			time.Sleep(100 * time.Millisecond)
			results <- fmt.Sprintf("处理完成 %d", id)
		}(i)
	}

	// 收集结果
	for i := 0; i < 5; i++ {
		result := <-results
		fmt.Printf("- %s\n", result)
	}

	concurrentTime := time.Since(startTime)
	fmt.Printf("并发处理时间: %v\n", concurrentTime)

	fmt.Println()
}

// createTestDocuments 创建测试文档
func createTestDocuments(count int) []*wordprocessingml.Document {
	var documents []*wordprocessingml.Document

	for i := 0; i < count; i++ {
		// 创建文档写入器
		w := writer.NewDocumentWriter()
		w.CreateNewDocument()

		// 添加内容
		w.AddParagraph(fmt.Sprintf("测试文档 %d", i+1), "Heading1")
		w.AddParagraph(fmt.Sprintf("这是测试文档 %d 的内容。", i+1), "Normal")
		w.AddParagraph("包含一些格式化文本。", "Normal")

		// 保存文档
		filename := fmt.Sprintf("test_doc_%d.docx", i+1)
		w.Save(filename)

		// 打开文档
		doc, err := wordprocessingml.Open(filename)
		if err != nil {
			log.Printf("打开文档 %s 失败: %v", filename, err)
			continue
		}

		documents = append(documents, doc)
	}

	return documents
}

// createProblematicDocument 创建有问题的文档
func createProblematicDocument() *wordprocessingml.Document {
	// 创建文档写入器
	w := writer.NewDocumentWriter()
	w.CreateNewDocument()

	// 添加一些有问题的内容
	w.AddParagraph("测试文档", "Heading1")
	w.AddParagraph("", "Normal") // 空段落
	w.AddParagraph("  包含重复空格  ", "Normal") // 重复空格
	w.AddParagraph("包含特殊字符", "Normal")

	// 保存文档
	w.Save("problematic_doc.docx")

	// 打开文档
	doc, err := wordprocessingml.Open("problematic_doc.docx")
	if err != nil {
		log.Printf("打开有问题的文档失败: %v", err)
		return nil
	}

	return doc
}

// createLargeDocument 创建大文档
func createLargeDocument() *wordprocessingml.Document {
	// 创建文档写入器
	w := writer.NewDocumentWriter()
	w.CreateNewDocument()

	// 添加大量内容
	for i := 0; i < 100; i++ {
		w.AddParagraph(fmt.Sprintf("段落 %d", i+1), "Normal")
		w.AddParagraph(fmt.Sprintf("这是第 %d 段的内容，包含一些测试文本。", i+1), "Normal")
	}

	// 保存文档
	w.Save("large_doc.docx")

	// 打开文档
	doc, err := wordprocessingml.Open("large_doc.docx")
	if err != nil {
		log.Printf("打开大文档失败: %v", err)
		return nil
	}

	return doc
}

// monitorProgress 监控进度
func monitorProgress(processor *wordprocessingml.BatchProcessor) {
	progressChan := processor.GetProgressChannel()
	
	for report := range progressChan {
		fmt.Printf("进度: %.1f%% (%d/%d) - %s\n", 
			report.Percentage, 
			report.ProcessedDocuments, 
			report.TotalDocuments,
			report.CurrentDocument)
	}
}

// monitorErrors 监控错误
func monitorErrors(processor *wordprocessingml.BatchProcessor) {
	errorChan := processor.GetErrorChannel()
	
	for batchError := range errorChan {
		fmt.Printf("错误: 文档 %s 操作 %s 失败: %v\n", 
			batchError.DocumentID, 
			batchError.Operation, 
			batchError.Error)
	}
} 