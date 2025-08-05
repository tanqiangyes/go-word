package examples

import (
	"fmt"
	"log"
	"runtime"
	"time"

	"github.com/tanqiangyes/go-word/pkg/utils"
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 性能演示示例
// 展示如何使用性能监控和优化工具
func DemoPerformance() {
	fmt.Println("=== Go OpenXML SDK 性能演示 ===")
	fmt.Println("本示例演示如何使用性能监控和优化工具")
	fmt.Println()

	// 显示系统信息
	utils.PrintSystemInfo()
	fmt.Println()

	// 显示性能优化建议
	utils.PrintPerformanceTips()
	fmt.Println()

	// 示例1: 基本性能监控
	fmt.Println("📊 示例1: 基本性能监控")
	demoBasicPerformanceMonitoring()
	fmt.Println()

	// 示例2: 内存分析
	fmt.Println("📊 示例2: 内存分析")
	demoMemoryAnalysis()
	fmt.Println()

	// 示例3: 批量处理性能
	fmt.Println("📊 示例3: 批量处理性能")
	demoBatchProcessingPerformance()
	fmt.Println()

	// 示例4: 错误处理和性能
	fmt.Println("📊 示例4: 错误处理和性能")
	demoErrorHandlingWithPerformance()
	fmt.Println()

	fmt.Println("✅ 性能演示完成")
}

// 基本性能监控演示
func demoBasicPerformanceMonitoring() {
	// 创建性能监控器
	monitor := utils.NewPerformanceMonitor(true)

	// 监控文档打开操作
	metrics := monitor.StartOperation("打开文档")
	
	// 模拟文档操作
	time.Sleep(100 * time.Millisecond)
	
	monitor.EndOperation(metrics)

	// 监控文本提取操作
	metrics = monitor.StartOperation("提取文本")
	
	// 模拟文本提取
	time.Sleep(50 * time.Millisecond)
	
	monitor.EndOperation(metrics)

	// 监控文档关闭操作
	metrics = monitor.StartOperation("关闭文档")
	
	// 模拟文档关闭
	time.Sleep(10 * time.Millisecond)
	
	monitor.EndOperation(metrics)

	// 打印性能摘要
	monitor.PrintSummary()

	// 分析性能并提供建议
	optimizer := utils.NewPerformanceOptimizer(monitor)
	suggestions := optimizer.AnalyzePerformance()
	
	fmt.Println("📋 性能分析建议:")
	for i, suggestion := range suggestions {
		fmt.Printf("  %d. %s\n", i+1, suggestion)
	}
}

// 内存分析演示
func demoMemoryAnalysis() {
	// 创建内存分析器
	profiler := utils.NewMemoryProfiler()

	// 记录初始状态
	profiler.TakeSnapshot("初始状态")

	// 模拟内存分配
	fmt.Println("  分配内存...")
	allocateMemory(10 * 1024 * 1024) // 10MB
	profiler.TakeSnapshot("分配10MB内存")

	// 模拟更多内存分配
	fmt.Println("  分配更多内存...")
	allocateMemory(20 * 1024 * 1024) // 20MB
	profiler.TakeSnapshot("分配20MB内存")

	// 模拟垃圾回收
	fmt.Println("  触发垃圾回收...")
	runtime.GC()
	profiler.TakeSnapshot("垃圾回收后")

	// 打印内存报告
	fmt.Println(profiler.GetMemoryReport())
}

// 批量处理性能演示
func demoBatchProcessingPerformance() {
	// 创建性能监控器
	monitor := utils.NewPerformanceMonitor(true)

	// 模拟批量处理多个文档
	documents := []string{"doc1.docx", "doc2.docx", "doc3.docx", "doc4.docx", "doc5.docx"}

	for i, docName := range documents {
		operation := fmt.Sprintf("处理文档 %d (%s)", i+1, docName)
		metrics := monitor.StartOperation(operation)

		// 模拟文档处理
		time.Sleep(time.Duration(100+i*20) * time.Millisecond)

		monitor.EndOperation(metrics)
	}

	// 打印性能摘要
	monitor.PrintSummary()

	// 分析批量处理性能
	optimizer := utils.NewPerformanceOptimizer(monitor)
	suggestions := optimizer.AnalyzePerformance()
	
	fmt.Println("📋 批量处理优化建议:")
	for i, suggestion := range suggestions {
		fmt.Printf("  %d. %s\n", i+1, suggestion)
	}
}

// 错误处理和性能演示
func demoErrorHandlingWithPerformance() {
	// 演示不同类型的错误处理
	fmt.Println("  测试文档错误处理...")
	
	// 模拟文档不存在错误
	err := simulateDocumentError("not_found")
	if err != nil {
		fmt.Printf("  ❌ 文档错误: %s\n", utils.GetUserFriendlyMessage(err))
	}

	// 模拟解析错误
	err = simulateParseError()
	if err != nil {
		fmt.Printf("  ❌ 解析错误: %s\n", utils.GetUserFriendlyMessage(err))
	}

	// 模拟I/O错误
	err = simulateIOError()
	if err != nil {
		fmt.Printf("  ❌ I/O错误: %s\n", utils.GetUserFriendlyMessage(err))
	}

	// 演示错误上下文
	fmt.Println("  测试错误上下文...")
	contextErr := utils.AddErrorContext(err, map[string]interface{}{
		"operation": "文档处理",
		"filename":  "test.docx",
		"timestamp": time.Now(),
	})
	fmt.Printf("  📝 带上下文的错误: %v\n", contextErr)
}

// 辅助函数

// allocateMemory 分配指定大小的内存用于测试
func allocateMemory(size int) {
	// 分配内存但不保留引用，让垃圾回收器可以回收
	data := make([]byte, size)
	
	// 使用数据避免编译器优化
	for i := 0; i < len(data); i += 1024 {
		data[i] = byte(i % 256)
	}
	
	// 模拟一些处理时间
	time.Sleep(10 * time.Millisecond)
}

// simulateDocumentError 模拟文档错误
func simulateDocumentError(errorType string) error {
	switch errorType {
	case "not_found":
		return utils.NewDocumentError("document not found", nil)
	case "corrupted":
		return utils.NewDocumentError("document is corrupted", nil)
	case "unsupported":
		return utils.NewDocumentError("unsupported document format", nil)
	default:
		return utils.NewDocumentError("unknown document error", nil)
	}
}

// simulateParseError 模拟解析错误
func simulateParseError() error {
	return utils.NewParseError("invalid XML structure", nil, 10, 25)
}

// simulateIOError 模拟I/O错误
func simulateIOError() error {
	return utils.NewIOError("permission denied", nil, "/path/to/document.docx", "read")
}

// 实际文档处理示例

// processDocumentWithMonitoring 使用性能监控处理文档
func processDocumentWithMonitoring(filename string) error {
	return utils.MeasureOperation("处理文档", func() error {
		// 打开文档
		doc, err := wordprocessingml.Open(filename)
		if err != nil {
			return utils.WrapError(err, "无法打开文档")
		}
		defer doc.Close()

		// 获取文本
		text, err := doc.GetText()
		if err != nil {
			return utils.WrapError(err, "无法获取文档文本")
		}

		// 获取段落
		paragraphs, err := doc.GetParagraphs()
		if err != nil {
			return utils.WrapError(err, "无法获取段落")
		}

		// 获取表格
		tables, err := doc.GetTables()
		if err != nil {
			return utils.WrapError(err, "无法获取表格")
		}

		// 打印统计信息
		fmt.Printf("  📄 文档统计:\n")
		fmt.Printf("    文本长度: %d 字符\n", len(text))
		fmt.Printf("    段落数量: %d\n", len(paragraphs))
		fmt.Printf("    表格数量: %d\n", len(tables))

		return nil
	})
}

// processDocumentWithMemoryProfiling 使用内存分析处理文档
func processDocumentWithMemoryProfiling(filename string) error {
	return utils.MeasureOperationWithMemory("处理文档(内存分析)", func() error {
		// 打开文档
		doc, err := wordprocessingml.Open(filename)
		if err != nil {
			return utils.WrapError(err, "无法打开文档")
		}
		defer doc.Close()

		// 获取文本
		text, err := doc.GetText()
		if err != nil {
			return utils.WrapError(err, "无法获取文档文本")
		}

		// 模拟一些处理
		fmt.Printf("  📄 处理了 %d 字符的文本\n", len(text))

		return nil
	})
}

// batchProcessDocuments 批量处理文档
func batchProcessDocuments(filenames []string) error {
	monitor := utils.NewPerformanceMonitor(true)

	for i, filename := range filenames {
		operation := fmt.Sprintf("处理文档 %d/%d", i+1, len(filenames))
		metrics := monitor.StartOperation(operation)

		// 处理文档
		err := processDocumentWithMonitoring(filename)
		if err != nil {
			log.Printf("处理文档 %s 时出错: %v", filename, err)
			// 继续处理其他文档
		}

		monitor.EndOperation(metrics)
	}

	// 打印批量处理摘要
	monitor.PrintSummary()

	// 分析批量处理性能
	optimizer := utils.NewPerformanceOptimizer(monitor)
	suggestions := optimizer.AnalyzePerformance()
	
	fmt.Println("📋 批量处理优化建议:")
	for i, suggestion := range suggestions {
		fmt.Printf("  %d. %s\n", i+1, suggestion)
	}

	return nil
} 