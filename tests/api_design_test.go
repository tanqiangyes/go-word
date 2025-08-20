package tests

import (
	"testing"

	"github.com/tanqiangyes/go-word/pkg/utils"
	"github.com/tanqiangyes/go-word/pkg/word"
)

func TestDocumentBuilder(t *testing.T) {
	// Test basic document builder
	builder := word.NewDocumentBuilder()

	doc, err := builder.
		WithTitle("测试文档").
		WithAuthor("测试作者").
		WithLanguage("zh-CN").
		WithDefaultFont("Microsoft YaHei", 12).
		WithMargins(72, 72, 72, 72).
		WithPageSize(595, 842, "portrait").
		WithProtection(word.ReadOnlyProtection, "password123").
		WithValidation(true, false, false).
		Build()

	if err != nil {
		t.Errorf("文档构建失败: %v", err)
	}

	if doc == nil {
		t.Error("文档不应该为 nil")
	}
}

func TestParagraphBuilder(t *testing.T) {
	// Test paragraph builder
	builder := word.NewParagraphBuilder()

	paragraph := builder.
		WithText("这是一个测试段落").
		WithStyle("Normal").
		WithComment("测试用户", "这是一个测试批注").
		Build()

	if paragraph.Text != "这是一个测试段落" {
		t.Errorf("段落文本不匹配，期望: '这是一个测试段落', 实际: '%s'", paragraph.Text)
	}

	if paragraph.Style != "Normal" {
		t.Errorf("段落样式不匹配，期望: 'Normal', 实际: '%s'", paragraph.Style)
	}

	if !paragraph.HasComment {
		t.Error("段落应该有批注")
	}

	if paragraph.CommentID == "" {
		t.Error("段落应该有批注ID")
	}
}

func TestTableBuilder(t *testing.T) {
	// Test table builder
	builder := word.NewTableBuilder()

	table := builder.
		WithHeaders("姓名", "年龄", "职业").
		WithRows(
			[]string{"张三", "25", "工程师"},
			[]string{"李四", "30", "设计师"},
		).
		WithStyle("TableGrid").
		Build()

	if table.Columns != 3 {
		t.Errorf("表格列数不匹配，期望: 3, 实际: %d", table.Columns)
	}

	if len(table.Rows) != 3 {
		t.Errorf("表格行数不匹配，期望: 3, 实际: %d", len(table.Rows))
	}

	// Check header row
	if len(table.Rows) > 0 {
		headerRow := table.Rows[0]
		if len(headerRow.Cells) != 3 {
			t.Errorf("表头单元格数不匹配，期望: 3, 实际: %d", len(headerRow.Cells))
		}

		if headerRow.Cells[0].Text != "姓名" {
			t.Errorf("表头内容不匹配，期望: '姓名', 实际: '%s'", headerRow.Cells[0].Text)
		}
	}
}

func TestErrorHandling(t *testing.T) {
	// Test structured error creation
	err := utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "文档未找到")

	if err.Code != utils.ErrDocumentNotFound {
		t.Errorf("错误代码不匹配，期望: %s, 实际: %s", utils.ErrDocumentNotFound, err.Code)
	}

	if err.Message != "文档未找到" {
		t.Errorf("错误消息不匹配，期望: '文档未找到', 实际: '%s'", err.Message)
	}

	// Test error with context
	err = err.WithContext("file_path", "/path/to/document.docx")

	if err.Context["file_path"] != "/path/to/document.docx" {
		t.Error("错误上下文不匹配")
	}

	// Test error with severity
	err = err.WithSeverity(utils.SeverityWarning)

	if err.Severity != utils.SeverityWarning {
		t.Errorf("错误严重性不匹配，期望: %d, 实际: %d", utils.SeverityWarning, err.Severity)
	}
}

func TestErrorCodeChecking(t *testing.T) {
	// Test error code checking
	err := utils.NewStructuredDocumentError(utils.ErrDocumentCorrupted, "文档损坏")

	if !utils.IsErrorCode(err, utils.ErrDocumentCorrupted) {
		t.Error("错误代码检查失败")
	}

	if utils.IsErrorCode(err, utils.ErrDocumentNotFound) {
		t.Error("错误代码检查应该失败")
	}

	// Test getting error code
	code := utils.GetErrorCode(err)
	if code != utils.ErrDocumentCorrupted {
		t.Errorf("获取错误代码失败，期望: %s, 实际: %s", utils.ErrDocumentCorrupted, code)
	}
}

func TestErrorSeverity(t *testing.T) {
	// Test error severity
	err := utils.NewStructuredDocumentError(utils.ErrDocumentProtected, "文档受保护")

	severity := utils.GetErrorSeverity(err)
	if severity != utils.SeverityError {
		t.Errorf("错误严重性不匹配，期望: %d, 实际: %d", utils.SeverityError, severity)
	}

	// Test with different severity
	err = err.WithSeverity(utils.SeverityCritical)
	severity = utils.GetErrorSeverity(err)
	if severity != utils.SeverityCritical {
		t.Errorf("错误严重性不匹配，期望: %d, 实际: %d", utils.SeverityCritical, severity)
	}
}

func TestErrorContext(t *testing.T) {
	// Test error context
	ctx := utils.NewErrorContext().
		WithDocumentPath("/path/to/document.docx").
		WithOperation("read_document").
		WithParameter("timeout", 30)

	if ctx.DocumentPath != "/path/to/document.docx" {
		t.Errorf("文档路径不匹配，期望: '/path/to/document.docx', 实际: '%s'", ctx.DocumentPath)
	}

	if ctx.Operation != "read_document" {
		t.Errorf("操作不匹配，期望: 'read_document', 实际: '%s'", ctx.Operation)
	}

	if ctx.Parameters["timeout"] != 30 {
		t.Error("参数不匹配")
	}
}

func TestErrorHandler(t *testing.T) {
	// Test error handler
	handler := utils.NewErrorHandler()

	// Register a custom handler
	handler.RegisterHandler(utils.ErrDocumentNotFound, func(err error, ctx *utils.ErrorContext) error {
		// Custom handling for document not found
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "文档未找到，已尝试自动恢复")
	})

	// Create an error
	err := utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "文档未找到")
	ctx := utils.NewErrorContext().WithDocumentPath("/test/document.docx")

	// Handle the error
	handledErr := handler.HandleError(err, ctx)

	if handledErr == nil {
		t.Error("错误处理应该返回错误")
	}

	// Check if the error was handled
	if utils.GetErrorCode(handledErr) != utils.ErrDocumentNotFound {
		t.Errorf("错误代码不匹配，期望: %s, 实际: %s", utils.ErrDocumentNotFound, utils.GetErrorCode(handledErr))
	}
}

func TestErrorRecovery(t *testing.T) {
	// Test error recovery
	handler := utils.NewErrorHandler()
	recovery := utils.NewErrorRecovery(handler)

	// Create an error
	err := utils.NewStructuredDocumentError(utils.ErrDocumentCorrupted, "文档损坏")
	ctx := utils.NewErrorContext().WithDocumentPath("/test/corrupted.docx")

	// Attempt recovery
	recoveredErr := recovery.RecoverFromError(err, ctx)

	if recoveredErr == nil {
		t.Error("错误恢复应该返回错误")
	}

	// Check if recovery was attempted
	if utils.GetErrorCode(recoveredErr) != utils.ErrDocumentCorrupted {
		t.Errorf("错误代码不匹配，期望: %s, 实际: %s", utils.ErrDocumentCorrupted, utils.GetErrorCode(recoveredErr))
	}
}

func TestErrorMetrics(t *testing.T) {
	// Test error metrics
	metrics := utils.NewErrorMetrics()

	// Record some errors
	metrics.RecordError(utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "文档未找到"))
	metrics.RecordError(utils.NewStructuredDocumentError(utils.ErrDocumentCorrupted, "文档损坏"))
	metrics.RecordError(utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "另一个文档未找到"))

	// Get metrics
	metricData := metrics.GetMetrics()

	if metricData["total_errors"] != 3 {
		t.Errorf("总错误数不匹配，期望: 3, 实际: %v", metricData["total_errors"])
	}

	errorCodes := metricData["error_codes"].(map[utils.ErrorCode]int)
	if errorCodes[utils.ErrDocumentNotFound] != 2 {
		t.Errorf("文档未找到错误数不匹配，期望: 2, 实际: %d", errorCodes[utils.ErrDocumentNotFound])
	}

	if errorCodes[utils.ErrDocumentCorrupted] != 1 {
		t.Errorf("文档损坏错误数不匹配，期望: 1, 实际: %d", errorCodes[utils.ErrDocumentCorrupted])
	}
}

func TestConfiguration(t *testing.T) {
	// Test configuration manager
	config := utils.NewConfigManager()

	// Test string configuration
	config.SetConfig("language", "zh-CN")
	if config.GetString("language") != "zh-CN" {
		t.Errorf("语言配置不匹配，期望: 'zh-CN', 实际: '%s'", config.GetString("language"))
	}

	// Test integer configuration
	config.SetConfig("timeout", 60)
	if config.GetInt("timeout") != 60 {
		t.Errorf("超时配置不匹配，期望: 60, 实际: %d", config.GetInt("timeout"))
	}

	// Test boolean configuration
	config.SetConfig("debug_mode", true)
	if !config.GetBool("debug_mode") {
		t.Error("调试模式配置不匹配")
	}

	// Test int64 configuration
	config.SetConfig("max_file_size", int64(100*1024*1024))
	if config.GetInt64("max_file_size") != 100*1024*1024 {
		t.Errorf("最大文件大小配置不匹配")
	}
}

func TestConfigurationValidation(t *testing.T) {
	// Test configuration validation
	config := utils.NewConfigManager().GetConfig()
	validator := utils.NewConfigValidator(config)

	errors := validator.Validate()

	if len(errors) > 0 {
		t.Errorf("配置验证失败: %v", errors)
	}

	// Test with invalid configuration
	config.Language = ""
	config.MaxFileSize = -1

	errors = validator.Validate()

	if len(errors) == 0 {
		t.Error("配置验证应该失败")
	}

	// Check specific errors
	foundLanguageError := false
	foundFileSizeError := false

	for _, err := range errors {
		if err == "language cannot be empty" {
			foundLanguageError = true
		}
		if err == "max_file_size must be positive" {
			foundFileSizeError = true
		}
	}

	if !foundLanguageError {
		t.Error("应该发现语言配置错误")
	}

	if !foundFileSizeError {
		t.Error("应该发现文件大小配置错误")
	}
}

func TestLogging(t *testing.T) {
	// Test logger creation
	logger := utils.NewLogger(utils.LogLevelInfo, nil)

	if logger == nil {
		t.Error("日志记录器不应该为 nil")
	}

	// Test log level setting
	logger.SetLevel(utils.LogLevelDebug)

	// Test formatter
	formatter := utils.NewDefaultFormatter(true, true, true)
	logger.SetFormatter(formatter)

	// Test handler
	consoleHandler := utils.NewConsoleHandler(false)
	logger.AddHandler(consoleHandler)

	// Test logging
	logger.Info("测试信息日志")
	logger.Warning("测试警告日志")
	logger.Error("测试错误日志")
}

func TestLogManager(t *testing.T) {
	// Test log manager
	manager := utils.NewLogManager()

	// Get logger
	logger := manager.GetLogger("test")

	if logger == nil {
		t.Error("日志记录器不应该为 nil")
	}

	// Test logging
	logger.Info("测试日志管理器")

	// Close manager
	manager.Close()
}

func TestPerformanceLogger(t *testing.T) {
	// Test performance logger
	logger := utils.NewLogger(utils.LogLevelInfo, nil)
	perfLogger := utils.NewPerformanceLogger(logger)

	// Test operation logging
	err := perfLogger.LogOperation("测试操作", func() error {
		// Simulate some work
		return nil
	})

	if err != nil {
		t.Errorf("性能日志记录失败: %v", err)
	}
}

func TestStructuredLogger(t *testing.T) {
	// Test structured logger
	logger := utils.NewLogger(utils.LogLevelInfo, nil)
	structuredLogger := utils.NewStructuredLogger(logger)

	// Test with fields
	structuredLogger.
		WithField("user_id", "123").
		WithField("operation", "read_document").
		WithFields(map[string]interface{}{
			"file_path": "/test/document.docx",
			"file_size": 1024,
		}).
		Info("结构化日志测试")
}

func TestFluentDocument(t *testing.T) {
	// Test fluent document interface
	doc := &word.Document{}
	fluentDoc := word.NewFluentDocument(doc)

	// Test paragraph builder
	paragraphBuilder := fluentDoc.AddParagraph()
	paragraph := paragraphBuilder.
		WithText("测试段落").
		WithStyle("Normal").
		Build()

	if paragraph.Text != "测试段落" {
		t.Errorf("段落文本不匹配，期望: '测试段落', 实际: '%s'", paragraph.Text)
	}

	// Test table builder
	tableBuilder := fluentDoc.AddTable()
	table := tableBuilder.
		WithHeaders("列1", "列2").
		WithRows([]string{"数据1", "数据2"}).
		Build()

	if table.Columns != 2 {
		t.Errorf("表格列数不匹配，期望: 2, 实际: %d", table.Columns)
	}
}
