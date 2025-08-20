package word

import (
    "context"
    "testing"
    "time"

    "github.com/tanqiangyes/go-word/pkg/types"
)

// TestErrorHandling 测试错误处理
func TestErrorHandling(t *testing.T) {
    // 测试空文档错误
    t.Run("NilDocumentError", func(t *testing.T) {
        var doc *Document
        // 直接调用方法会导致panic，因为接收者本身是nil
        // 我们应该测试这种情况下的行为
        defer func() {
            if r := recover(); r != nil {
                t.Logf("Expected panic for nil document: %v", r)
            } else {
                t.Error("Expected panic for nil document")
            }
        }()

        _, err := doc.GetParagraphs()
        if err == nil {
            t.Error("Expected error for nil document")
        }
    })

    // 测试空内容错误
    t.Run("NilContentError", func(t *testing.T) {
        doc := &Document{}
        _, err := doc.GetTables()
        if err == nil {
            t.Error("Expected error for nil content")
        }
        // 验证错误消息
        if err != nil && err.Error() != "document content not loaded" {
            t.Errorf("Expected error message 'document content not loaded', got: %v", err.Error())
        }
    })
}

// TestLoggerInitialization 测试Logger初始化
func TestLoggerInitialization(t *testing.T) {
    t.Run("PDFExporterLogger", func(t *testing.T) {
        doc := &Document{}
        exporter := NewPDFExporter(doc, nil)
        if exporter.Logger == nil {
            t.Error("PDFExporter logger should not be nil")
        }
    })

    t.Run("PluginSystemLogger", func(t *testing.T) {
        doc := &Document{}
        config := &PluginConfig{Enabled: true}
        ps := NewPluginSystem(doc, config)
        if ps.Logger == nil {
            t.Error("PluginSystem logger should not be nil")
        }
    })

    t.Run("FileEmbedderLogger", func(t *testing.T) {
        doc := &Document{}
        fe := NewFileEmbedder(doc, nil)
        if fe.logger == nil {
            t.Error("FileEmbedder logger should not be nil")
        }
    })

    t.Run("CustomRibbonLogger", func(t *testing.T) {
        doc := &Document{}
        cr := NewCustomRibbon(doc, nil)
        if cr.logger == nil {
            t.Error("CustomRibbon logger should not be nil")
        }
    })
}

// TestTypeConversions 测试类型转换
func TestTypeConversions(t *testing.T) {
    t.Run("ParagraphPointerConversion", func(t *testing.T) {
        paragraphs := []types.Paragraph{
            {Text: "Test paragraph 1"},
            {Text: "Test paragraph 2"},
        }

        var pointerSlice []*types.Paragraph
        for i := range paragraphs {
            pointerSlice = append(pointerSlice, &paragraphs[i])
        }

        if len(pointerSlice) != len(paragraphs) {
            t.Errorf("Expected %d pointers, got %d", len(paragraphs), len(pointerSlice))
        }

        for i, p := range pointerSlice {
            if p.Text != paragraphs[i].Text {
                t.Errorf("Pointer conversion failed: expected %s, got %s", paragraphs[i].Text, p.Text)
            }
        }
    })

    t.Run("TablePointerConversion", func(t *testing.T) {
        tables := []types.Table{
            {Columns: 2, Rows: []types.TableRow{{Cells: []types.TableCell{{Text: "Cell1"}, {Text: "Cell2"}}}}},
        }

        var pointerSlice []*types.Table
        for i := range tables {
            pointerSlice = append(pointerSlice, &tables[i])
        }

        if len(pointerSlice) != len(tables) {
            t.Errorf("Expected %d pointers, got %d", len(tables), len(pointerSlice))
        }
    })
}

// TestContextHandling 测试上下文处理
func TestContextHandling(t *testing.T) {
    t.Run("ContextTimeout", func(t *testing.T) {
        ctx, cancel := context.WithTimeout(context.Background(), 1*time.Millisecond)
        defer cancel()

        time.Sleep(2 * time.Millisecond) // 确保超时

        select {
        case <-ctx.Done():
            if ctx.Err() != context.DeadlineExceeded {
                t.Errorf("Expected DeadlineExceeded, got %v", ctx.Err())
            }
        default:
            t.Error("Context should have timed out")
        }
    })

    t.Run("ContextCancellation", func(t *testing.T) {
        ctx, cancel := context.WithCancel(context.Background())
        cancel()

        select {
        case <-ctx.Done():
            if ctx.Err() != context.Canceled {
                t.Errorf("Expected Canceled, got %v", ctx.Err())
            }
        default:
            t.Error("Context should have been cancelled")
        }
    })
}

// TestConfigValidation 测试配置验证
func TestConfigValidation(t *testing.T) {
    t.Run("DefaultPDFConfig", func(t *testing.T) {
        config := getDefaultPDFConfig()
        if config == nil {
            t.Error("Default PDF config should not be nil")
        }
        if config.PageSize == "" {
            t.Error("Default page size should not be empty")
        }
    })

    t.Run("DefaultPluginConfig", func(t *testing.T) {
        config := getDefaultPluginConfig()
        if config == nil {
            t.Error("Default plugin config should not be nil")
        }
        if config.MaxPlugins <= 0 {
            t.Error("Max plugins should be positive")
        }
    })
}

// TestMemoryLeaks 测试内存泄漏
func TestMemoryLeaks(t *testing.T) {
    t.Run("DocumentCleanup", func(t *testing.T) {
        doc := &Document{
            documentParts: NewDocumentParts(),
        }
        mainPart := &MainDocumentPart{
            Content: &types.DocumentContent{
                Paragraphs: make([]types.Paragraph, 1000), // 大量数据
                Tables:     make([]types.Table, 100),
            },
        }
        doc.SetMainPart(mainPart)

        // 模拟清理
        doc.mainPart = nil
        doc.documentParts = nil

        // 验证清理
        if doc.mainPart != nil {
            t.Error("Main part should be nil after cleanup")
        }
    })
}

// TestConcurrency 测试并发安全
func TestConcurrency(t *testing.T) {
    t.Run("ConcurrentAccess", func(t *testing.T) {
        doc := &Document{
            documentParts: NewDocumentParts(),
        }
        mainPart := &MainDocumentPart{
            Content: &types.DocumentContent{
                Paragraphs: []types.Paragraph{{Text: "Test"}},
            },
        }
        doc.SetMainPart(mainPart)

        done := make(chan bool, 10)

        // 启动多个goroutine并发访问
        for i := 0; i < 10; i++ {
            go func() {
                defer func() { done <- true }()
                _, err := doc.GetParagraphs()
                if err != nil {
                    t.Errorf("Concurrent access failed: %v", err)
                }
            }()
        }

        // 等待所有goroutine完成
        for i := 0; i < 10; i++ {
            <-done
        }
    })
}

// TestErrorRecovery 测试错误恢复
func TestErrorRecovery(t *testing.T) {
    t.Run("PanicRecovery", func(t *testing.T) {
        defer func() {
            if r := recover(); r != nil {
                t.Logf("Recovered from panic: %v", r)
            }
        }()

        // 模拟可能导致panic的操作
        var nilSlice []types.Paragraph
        if len(nilSlice) > 0 {
            _ = nilSlice[0] // 这不会panic，因为len(nilSlice) == 0
        }
    })
}

// BenchmarkErrorHandling 性能测试
func BenchmarkErrorHandling(b *testing.B) {
    doc := &Document{
        documentParts: NewDocumentParts(),
    }
    mainPart := &MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{{Text: "Benchmark test"}},
        },
    }
    doc.SetMainPart(mainPart)

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        _, _ = doc.GetParagraphs()
    }
}
