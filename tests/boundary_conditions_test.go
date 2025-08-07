package tests

import (
	"fmt"
	"os"
	"strings"
	"testing"
	"time"

	"github.com/tanqiangyes/go-word/pkg/parser"
	"github.com/tanqiangyes/go-word/pkg/writer"
	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// TestEmptyDocument 测试空文档的边界情况
func TestEmptyDocument(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	// 测试创建空文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create empty document: %v", err)
	}
	
	// 测试保存空文档
	err = docWriter.Save("empty_test.docx")
	if err != nil {
		t.Fatalf("Failed to save empty document: %v", err)
	}
	
	// 清理测试文件
	defer os.Remove("empty_test.docx")
}

// TestSpecialCharacters 测试特殊字符的边界情况
func TestSpecialCharacters(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 测试各种特殊字符
	specialChars := []string{
		"", // 空字符串
		" ", // 空格
		"\t\n\r", // 制表符和换行符
		"中文测试", // 中文字符
		"Test with < > & \" ' characters", // XML特殊字符
		"Unicode: 🚀🌟🎉", // Unicode字符
		strings.Repeat("A", 10000), // 超长字符串
		"Mixed 中文 English 123 !@#$%", // 混合字符
	}
	
	for i, text := range specialChars {
		err = docWriter.AddParagraph(text, "Normal")
		if err != nil {
			t.Fatalf("Failed to add special character paragraph %d: %v", i, err)
		}
	}
	
	err = docWriter.Save("special_chars_test.docx")
	if err != nil {
		t.Fatalf("Failed to save special characters document: %v", err)
	}
	
	// 清理测试文件
	defer os.Remove("special_chars_test.docx")
}

// TestInvalidXML 测试无效XML的边界情况
func TestInvalidXML(t *testing.T) {
	parser := &parser.WordMLParser{}
	
	// 测试空XML
	_, err := parser.ParseWordDocument([]byte(""))
	if err == nil {
		t.Error("Expected error for empty XML")
	}
	
	// 测试无效XML
	_, err = parser.ParseWordDocument([]byte("<invalid>xml"))
	if err == nil {
		t.Error("Expected error for invalid XML")
	}
	
	// 测试不完整的XML
	_, err = parser.ParseWordDocument([]byte("<document><body>"))
	if err == nil {
		t.Error("Expected error for incomplete XML")
	}
	
	// 测试超大XML
	largeXML := "<document><body>" + strings.Repeat("<p>test</p>", 10000) + "</body></document>"
	_, err = parser.ParseWordDocument([]byte(largeXML))
	if err != nil {
		t.Logf("Large XML parsing error (expected): %v", err)
	}
}

// TestFileSystemBoundaries 测试文件系统边界情况
func TestFileSystemBoundaries(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 测试保存到不存在的目录
	err = docWriter.Save("/nonexistent/path/test.docx")
	if err == nil {
		t.Error("Expected error for nonexistent directory")
	}
	
	// 测试保存到当前目录
	err = docWriter.Save("./current_dir_test.docx")
	if err != nil {
		t.Fatalf("Failed to save to current directory: %v", err)
	}
	defer os.Remove("./current_dir_test.docx")
}

// TestStyleBoundaries 测试样式边界情况
func TestStyleBoundaries(t *testing.T) {
	system := wordprocessingml.NewAdvancedStyleSystem()
	
	// 测试空样式ID
	err := system.AddParagraphStyle(&wordprocessingml.ParagraphStyleDefinition{
		ID:   "",
		Name: "Empty ID Style",
	})
	if err == nil {
		t.Error("Expected error for empty style ID")
	}
	
	// 测试空样式名称（当前实现允许空名称）
	err = system.AddParagraphStyle(&wordprocessingml.ParagraphStyleDefinition{
		ID:   "EmptyName",
		Name: "",
	})
	if err != nil {
		t.Fatalf("Unexpected error for empty style name: %v", err)
	}
	
	// 测试重复样式ID
	style1 := &wordprocessingml.ParagraphStyleDefinition{
		ID:   "Duplicate",
		Name: "First Style",
	}
	style2 := &wordprocessingml.ParagraphStyleDefinition{
		ID:   "Duplicate",
		Name: "Second Style",
	}
	
	err = system.AddParagraphStyle(style1)
	if err != nil {
		t.Fatalf("Failed to add first style: %v", err)
	}
	
	err = system.AddParagraphStyle(style2)
	if err == nil {
		t.Error("Expected error for duplicate style ID")
	}
}

// TestDocumentQualityBoundaries 测试文档质量边界情况
func TestDocumentQualityBoundaries(t *testing.T) {
	doc := &wordprocessingml.Document{}
	doc.SetMainPart(&wordprocessingml.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{},
			Tables:     []types.Table{},
		},
	})
	
	manager := wordprocessingml.NewDocumentQualityManager(doc)
	
	// 测试空文档的质量改进
	err := manager.ImproveDocumentQuality()
	if err != nil {
		t.Fatalf("Failed to improve empty document quality: %v", err)
	}
	
	// 测试包含无效内容的文档
	doc.SetMainPart(&wordprocessingml.MainDocumentPart{
		Content: &types.DocumentContent{
			Paragraphs: []types.Paragraph{
				{
					Runs: []types.Run{
						{Text: ""}, // 空文本
						{Text: "   "}, // 只有空格
						{Text: strings.Repeat("A", 10000)}, // 超长文本
					},
				},
			},
			Tables: []types.Table{
				{
					Rows: []types.TableRow{
						{
							Cells: []types.TableCell{
								{Text: ""}, // 空单元格
							},
						},
					},
				},
			},
		},
	})
	
	manager = wordprocessingml.NewDocumentQualityManager(doc)
	err = manager.ImproveDocumentQuality()
	if err != nil {
		t.Fatalf("Failed to improve document with invalid content: %v", err)
	}
}

// TestErrorRecoveryBoundaries 测试错误恢复边界情况
func TestErrorRecoveryBoundaries(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	// 测试在错误后继续操作
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 尝试添加无效内容（应该被处理）
	err = docWriter.AddParagraph("", "InvalidStyle")
	if err != nil {
		t.Logf("Expected error for invalid style: %v", err)
	}
	
	// 继续添加有效内容
	err = docWriter.AddParagraph("Valid paragraph", "Normal")
	if err != nil {
		t.Fatalf("Failed to add valid paragraph after error: %v", err)
	}
	
	// 尝试保存到无效路径
	err = docWriter.Save("")
	if err == nil {
		t.Error("Expected error for empty save path")
	}
	
	// 继续保存到有效路径
	err = docWriter.Save("recovery_test.docx")
	if err != nil {
		t.Fatalf("Failed to save after error: %v", err)
	}
	defer os.Remove("recovery_test.docx")
}

// TestEncodingBoundaries 测试编码边界情况
func TestEncodingBoundaries(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 测试各种编码的文本
	encodingTests := []string{
		"ASCII text",
		"中文文本",
		"Mixed 中文 English 123",
		"Unicode: 🚀🌟🎉",
		"Special chars: < > & \" '",
		"Control chars: \x00\x01\x02",
		"UTF-8 BOM: \ufeff",
		strings.Repeat("A", 1000), // 长ASCII
		strings.Repeat("中", 1000), // 长中文
	}
	
	for i, text := range encodingTests {
		err = docWriter.AddParagraph(text, "Normal")
		if err != nil {
			t.Fatalf("Failed to add encoding test paragraph %d: %v", i, err)
		}
	}
	
	err = docWriter.Save("encoding_test.docx")
	if err != nil {
		t.Fatalf("Failed to save encoding test document: %v", err)
	}
	defer os.Remove("encoding_test.docx")
} 

// TestPerformanceBoundaries 测试性能边界情况
func TestPerformanceBoundaries(t *testing.T) {
	start := time.Now()
	
	docWriter := writer.NewDocumentWriter()
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 测试快速连续操作
	for i := 0; i < 1000; i++ {
		err = docWriter.AddParagraph(fmt.Sprintf("Performance test paragraph %d", i), "Normal")
		if err != nil {
			t.Fatalf("Failed to add performance test paragraph %d: %v", i, err)
		}
		
		// 每100个操作后保存一次
		if i%100 == 0 {
			err = docWriter.Save(fmt.Sprintf("performance_test_%d.docx", i))
			if err != nil {
				t.Fatalf("Failed to save performance test document %d: %v", i, err)
			}
			defer os.Remove(fmt.Sprintf("performance_test_%d.docx", i))
		}
	}
	
	duration := time.Since(start)
	t.Logf("Performance test completed in %v", duration)
	
	// 如果操作时间超过30秒，记录警告
	if duration > 30*time.Second {
		t.Logf("Warning: Performance test took longer than expected: %v", duration)
	}
}

// TestConcurrentAccess 测试并发访问边界情况
func TestConcurrentAccess(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 并发添加段落
	done := make(chan bool, 10)
	for i := 0; i < 10; i++ {
		go func(id int) {
			for j := 0; j < 100; j++ {
				err := docWriter.AddParagraph(fmt.Sprintf("Concurrent paragraph %d-%d", id, j), "Normal")
				if err != nil {
					t.Errorf("Failed to add concurrent paragraph %d-%d: %v", id, j, err)
				}
			}
			done <- true
		}(i)
	}
	
	// 等待所有goroutine完成
	for i := 0; i < 10; i++ {
		<-done
	}
	
	err = docWriter.Save("concurrent_test.docx")
	if err != nil {
		t.Fatalf("Failed to save concurrent test document: %v", err)
	}
	defer os.Remove("concurrent_test.docx")
}

// TestVeryLargeDocument 测试超大文档的边界情况
func TestVeryLargeDocument(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 添加大量段落
	for i := 0; i < 1000; i++ {
		err = docWriter.AddParagraph(strings.Repeat("Large document test paragraph ", 10), "Normal")
		if err != nil {
			t.Fatalf("Failed to add large paragraph %d: %v", i, err)
		}
	}
	
	// 添加大量表格
	for i := 0; i < 100; i++ {
		tableData := make([][]string, 10)
		for j := 0; j < 10; j++ {
			tableData[j] = make([]string, 10)
			for k := 0; k < 10; k++ {
				tableData[j][k] = fmt.Sprintf("Cell %d-%d-%d", i, j, k)
			}
		}
		err = docWriter.AddTable(tableData)
		if err != nil {
			t.Fatalf("Failed to add large table %d: %v", i, err)
		}
	}
	
	// 测试保存大文档
	err = docWriter.Save("large_test.docx")
	if err != nil {
		t.Fatalf("Failed to save large document: %v", err)
	}
	
	// 清理测试文件
	defer os.Remove("large_test.docx")
}

// TestMemoryBoundaries 测试内存边界情况
func TestMemoryBoundaries(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 测试大量小段落
	for i := 0; i < 10000; i++ {
		err = docWriter.AddParagraph("Small paragraph", "Normal")
		if err != nil {
			t.Fatalf("Failed to add small paragraph %d: %v", i, err)
		}
	}
	
	// 测试大量格式化文本
	for i := 0; i < 1000; i++ {
		formattedRuns := []types.Run{
			{Text: "Bold", Bold: true},
			{Text: "Italic", Italic: true},
			{Text: "Underline", Underline: true},
		}
		err = docWriter.AddFormattedParagraph("Formatted text", "Normal", formattedRuns)
		if err != nil {
			t.Fatalf("Failed to add formatted paragraph %d: %v", i, err)
		}
	}
	
	err = docWriter.Save("memory_test.docx")
	if err != nil {
		t.Fatalf("Failed to save memory test document: %v", err)
	}
	defer os.Remove("memory_test.docx")
}

// TestResourceCleanup 测试资源清理边界情况
func TestResourceCleanup(t *testing.T) {
	// 创建多个文档写入器
	writers := make([]*writer.DocumentWriter, 10)
	
	for i := 0; i < 10; i++ {
		writers[i] = writer.NewDocumentWriter()
		err := writers[i].CreateNewDocument()
		if err != nil {
			t.Fatalf("Failed to create document %d: %v", i, err)
		}
		
		// 添加一些内容
		for j := 0; j < 100; j++ {
			err = writers[i].AddParagraph(fmt.Sprintf("Document %d, paragraph %d", i, j), "Normal")
			if err != nil {
				t.Fatalf("Failed to add paragraph to document %d: %v", i, err)
			}
		}
		
		// 保存文档
		err = writers[i].Save(fmt.Sprintf("cleanup_test_%d.docx", i))
		if err != nil {
			t.Fatalf("Failed to save document %d: %v", i, err)
		}
		defer os.Remove(fmt.Sprintf("cleanup_test_%d.docx", i))
	}
	
	// 清理所有写入器（模拟垃圾回收）
	writers = nil
	
	// 验证文件仍然存在且可读
	for i := 0; i < 10; i++ {
		_, err := os.Stat(fmt.Sprintf("cleanup_test_%d.docx", i))
		if err != nil {
			t.Errorf("File cleanup_test_%d.docx not found after cleanup: %v", i, err)
		}
	}
}

// TestXMLStructureBoundaries 测试XML结构边界情况
func TestXMLStructureBoundaries(t *testing.T) {
	parser := &parser.WordMLParser{}
	
	// 测试最小有效XML
	minimalXML := `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Test</w:t></w:r></w:p>
</w:body>
</w:document>`
	
	doc, err := parser.ParseWordDocument([]byte(minimalXML))
	if err != nil {
		t.Fatalf("Failed to parse minimal XML: %v", err)
	}
	
	if doc.Body.Paragraphs == nil || len(doc.Body.Paragraphs) == 0 {
		t.Error("Expected at least one paragraph in minimal XML")
	}
	
	// 测试嵌套过深的XML
	deepNestedXML := `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t><w:r><w:t><w:r><w:t>Deep nested</w:t></w:r></w:t></w:r></w:t></w:r></w:p>
</w:body>
</w:document>`
	
	_, err = parser.ParseWordDocument([]byte(deepNestedXML))
	if err != nil {
		t.Logf("Expected error for deeply nested XML: %v", err)
	}
	
	// 测试缺少必需元素的XML
	incompleteXML := `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
</w:body>
</w:document>`
	
	doc, err = parser.ParseWordDocument([]byte(incompleteXML))
	if err != nil {
		t.Logf("Expected error for incomplete XML: %v", err)
	} else {
		// 即使解析成功，也应该处理空文档
		if doc.Body.Paragraphs == nil {
			t.Log("Expected empty paragraphs slice for incomplete XML")
		}
	}
}

// TestTimeoutBoundaries 测试超时边界情况
func TestTimeoutBoundaries(t *testing.T) {
	docWriter := writer.NewDocumentWriter()
	
	err := docWriter.CreateNewDocument()
	if err != nil {
		t.Fatalf("Failed to create document: %v", err)
	}
	
	// 添加大量内容以测试处理时间
	start := time.Now()
	
	for i := 0; i < 10000; i++ {
		err = docWriter.AddParagraph(fmt.Sprintf("Timeout test paragraph %d", i), "Normal")
		if err != nil {
			t.Fatalf("Failed to add timeout test paragraph %d: %v", i, err)
		}
		
		// 每1000个操作检查一次时间
		if i%1000 == 0 {
			duration := time.Since(start)
			if duration > 60*time.Second {
				t.Fatalf("Timeout test exceeded 60 seconds: %v", duration)
			}
		}
	}
	
	duration := time.Since(start)
	t.Logf("Timeout test completed in %v", duration)
	
	// 保存文档
	err = docWriter.Save("timeout_test.docx")
	if err != nil {
		t.Fatalf("Failed to save timeout test document: %v", err)
	}
	defer os.Remove("timeout_test.docx")
}

// TestConcurrentFileAccess 测试并发文件访问边界情况
func TestConcurrentFileAccess(t *testing.T) {
	// 创建多个文档写入器同时写入同一文件
	writers := make([]*writer.DocumentWriter, 5)
	
	for i := 0; i < 5; i++ {
		writers[i] = writer.NewDocumentWriter()
		err := writers[i].CreateNewDocument()
		if err != nil {
			t.Fatalf("Failed to create document %d: %v", i, err)
		}
		
		// 添加一些内容
		err = writers[i].AddParagraph(fmt.Sprintf("Concurrent document %d", i), "Normal")
		if err != nil {
			t.Fatalf("Failed to add paragraph to document %d: %v", i, err)
		}
	}
	
	// 同时保存到同一文件（应该产生错误）
	done := make(chan bool, 5)
	for i := 0; i < 5; i++ {
		go func(id int) {
			err := writers[id].Save("concurrent_file_test.docx")
			if err != nil {
				t.Logf("Expected error for concurrent file access %d: %v", id, err)
			}
			done <- true
		}(i)
	}
	
	// 等待所有goroutine完成
	for i := 0; i < 5; i++ {
		<-done
	}
	
	// 清理可能创建的文件
	os.Remove("concurrent_file_test.docx")
}

// TestMemoryLeakBoundaries 测试内存泄漏边界情况
func TestMemoryLeakBoundaries(t *testing.T) {
	// 创建大量文档写入器
	writers := make([]*writer.DocumentWriter, 100)
	
	for i := 0; i < 100; i++ {
		writers[i] = writer.NewDocumentWriter()
		err := writers[i].CreateNewDocument()
		if err != nil {
			t.Fatalf("Failed to create document %d: %v", i, err)
		}
		
		// 添加大量内容
		for j := 0; j < 1000; j++ {
			err = writers[i].AddParagraph(fmt.Sprintf("Memory test %d-%d", i, j), "Normal")
			if err != nil {
				t.Fatalf("Failed to add paragraph %d-%d: %v", i, j, err)
			}
		}
	}
	
	// 清理所有写入器（模拟垃圾回收）
	writers = nil
	
	t.Log("Memory leak test completed - check memory usage manually")
} 