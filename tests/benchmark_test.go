package tests

import (
    "bytes"
    "os"
    "testing"

    "github.com/tanqiangyes/go-word/pkg/opc"
    "github.com/tanqiangyes/go-word/pkg/parser"
    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

var testContent []byte

func init() {
    // 读取测试文档用于基准测试
    var err error
    testContent, err = os.ReadFile("test_document.xml")
    if err != nil {
        panic("Failed to read test document for benchmarks")
    }
}

// BenchmarkXMLParsing 测试XML解析性能
func BenchmarkXMLParsing(b *testing.B) {
    parser := &parser.WordMLParser{}

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        _, err := parser.ParseWordDocument(testContent)
        if err != nil {
            b.Fatalf("Failed to parse document: %v", err)
        }
    }
}

// BenchmarkTextExtraction 测试文本提取性能
func BenchmarkTextExtraction(b *testing.B) {
    parser := &parser.WordMLParser{}
    doc, err := parser.ParseWordDocument(testContent)
    if err != nil {
        b.Fatalf("Failed to parse document: %v", err)
    }

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        parser.ExtractText(doc)
    }
}

// BenchmarkParagraphExtraction 测试段落提取性能
func BenchmarkParagraphExtraction(b *testing.B) {
    parser := &parser.WordMLParser{}
    doc, err := parser.ParseWordDocument(testContent)
    if err != nil {
        b.Fatalf("Failed to parse document: %v", err)
    }

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        parser.ExtractParagraphs(doc)
    }
}

// BenchmarkTableExtraction 测试表格提取性能
func BenchmarkTableExtraction(b *testing.B) {
    parser := &parser.WordMLParser{}
    doc, err := parser.ParseWordDocument(testContent)
    if err != nil {
        b.Fatalf("Failed to parse document: %v", err)
    }

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        parser.ExtractTables(doc)
    }
}

// BenchmarkOPCContainer 测试OPC容器操作性能
func BenchmarkOPCContainer(b *testing.B) {
    // 创建一个简单的ZIP数据用于测试
    zipData := createTestZipData()

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        container, err := opc.OpenFromReader(bytes.NewReader(zipData))
        if err != nil {
            b.Fatalf("Failed to open container: %v", err)
        }

        _, err = container.ListParts()
        if err != nil {
            b.Fatalf("Failed to list parts: %v", err)
        }

        container.Close()
    }
}

// BenchmarkDocumentOpen 测试文档打开性能
func BenchmarkDocumentOpen(b *testing.B) {
    filename := "example.docx"

    // 检查文件是否存在
    if _, err := os.Stat(filename); os.IsNotExist(err) {
        b.Skip("Test file not found, skipping benchmark")
    }

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        doc, err := word.Open(filename)
        if err != nil {
            b.Fatalf("Failed to open document: %v", err)
        }

        _, err = doc.GetText()
        if err != nil {
            b.Fatalf("Failed to get text: %v", err)
        }

        doc.Close()
    }
}

// BenchmarkMemoryUsage 测试内存使用情况
func BenchmarkMemoryUsage(b *testing.B) {
    parser := &parser.WordMLParser{}

    // 运行多次以测试内存使用
    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        doc, err := parser.ParseWordDocument(testContent)
        if err != nil {
            b.Fatalf("Failed to parse document: %v", err)
        }

        // 执行各种操作
        parser.ExtractText(doc)
        parser.ExtractParagraphs(doc)
        parser.ExtractTables(doc)
    }
}

// 辅助函数：创建测试ZIP数据
func createTestZipData() []byte {
    // 这是一个简化的实现
    // 在实际测试中，应该创建更真实的ZIP文件
    return []byte("PK\x03\x04\x14\x00\x00\x00\x08\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00test.xml\x00\x00\x00")
}
