// Package utils provides utility functions and logging functionality
package utils

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"time"
)

// LogLevel represents the logging level
type LogLevel int

const (
	LogLevelDebug LogLevel = iota
	LogLevelInfo
	LogLevelWarning
	LogLevelError
	LogLevelCritical
)

// String returns the string representation of the log level
func (l LogLevel) String() string {
	switch l {
	case LogLevelDebug:
		return "DEBUG"
	case LogLevelInfo:
		return "INFO"
	case LogLevelWarning:
		return "WARNING"
	case LogLevelError:
		return "ERROR"
	case LogLevelCritical:
		return "CRITICAL"
	default:
		return "UNKNOWN"
	}
}

// Logger provides logging functionality
type Logger struct {
	level      LogLevel
	output     io.Writer
	formatter  LogFormatter
	handlers   []LogHandler
	context    map[string]interface{}
	callerInfo bool
}

// LogFormatter formats log messages
type LogFormatter interface {
	Format(level LogLevel, message string, context map[string]interface{}, caller string) string
}

// LogHandler handles log messages
type LogHandler interface {
	Handle(level LogLevel, message string, context map[string]interface{})
}

// DefaultFormatter provides default log formatting
type DefaultFormatter struct {
	includeTimestamp bool
	includeCaller    bool
	includeContext   bool
}

// NewDefaultFormatter creates a new default formatter
func NewDefaultFormatter(includeTimestamp, includeCaller, includeContext bool) *DefaultFormatter {
	return &DefaultFormatter{
		includeTimestamp: includeTimestamp,
		includeCaller:    includeCaller,
		includeContext:   includeContext,
	}
}

// Format formats a log message
func (f *DefaultFormatter) Format(level LogLevel, message string, context map[string]interface{}, caller string) string {
	var parts []string
	
	// Add timestamp
	if f.includeTimestamp {
		parts = append(parts, time.Now().Format("2006-01-02 15:04:05"))
	}
	
	// Add level
	parts = append(parts, fmt.Sprintf("[%s]", level.String()))
	
	// Add caller information
	if f.includeCaller && caller != "" {
		parts = append(parts, fmt.Sprintf("[%s]", caller))
	}
	
	// Add message
	parts = append(parts, message)
	
	// Add context
	if f.includeContext && len(context) > 0 {
		var contextParts []string
		for key, value := range context {
			contextParts = append(contextParts, fmt.Sprintf("%s=%v", key, value))
		}
		parts = append(parts, fmt.Sprintf("(%s)", strings.Join(contextParts, ", ")))
	}
	
	return strings.Join(parts, " ")
}

// FileHandler writes logs to a file
type FileHandler struct {
	file     *os.File
	writer   io.Writer
	rotation LogRotation
}

// LogRotation provides log file rotation functionality
type LogRotation struct {
	MaxSize    int64
	MaxAge     time.Duration
	MaxBackups int
}

// NewFileHandler creates a new file handler
func NewFileHandler(filename string, rotation LogRotation) (*FileHandler, error) {
	// Create directory if it doesn't exist
	dir := filepath.Dir(filename)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return nil, fmt.Errorf("failed to create log directory: %w", err)
	}
	
	// Open file
	file, err := os.OpenFile(filename, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err != nil {
		return nil, fmt.Errorf("failed to open log file: %w", err)
	}
	
	return &FileHandler{
		file:     file,
		writer:   file,
		rotation: rotation,
	}, nil
}

// Handle handles a log message
func (h *FileHandler) Handle(level LogLevel, message string, context map[string]interface{}) {
	// Check if rotation is needed
	if h.rotation.MaxSize > 0 {
		if info, err := h.file.Stat(); err == nil {
			if info.Size() >= h.rotation.MaxSize {
				h.rotate()
			}
		}
	}
	
	// Write log message
	fmt.Fprintln(h.writer, message)
}

// rotate rotates the log file
func (h *FileHandler) rotate() {
	// Close current file
	h.file.Close()
	
	// Rename current file
	oldName := h.file.Name()
	newName := oldName + "." + time.Now().Format("2006-01-02-15-04-05")
	os.Rename(oldName, newName)
	
	// Open new file
	file, err := os.OpenFile(oldName, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err == nil {
		h.file = file
		h.writer = file
	}
}

// Close closes the file handler
func (h *FileHandler) Close() error {
	return h.file.Close()
}

// ConsoleHandler writes logs to console
type ConsoleHandler struct {
	writer io.Writer
	color  bool
}

// NewConsoleHandler creates a new console handler
func NewConsoleHandler(color bool) *ConsoleHandler {
	return &ConsoleHandler{
		writer: os.Stdout,
		color:  color,
	}
}

// Handle handles a log message
func (h *ConsoleHandler) Handle(level LogLevel, message string, context map[string]interface{}) {
	if h.color {
		// Add color codes
		var colorCode string
		switch level {
		case LogLevelDebug:
			colorCode = "\033[36m" // Cyan
		case LogLevelInfo:
			colorCode = "\033[32m" // Green
		case LogLevelWarning:
			colorCode = "\033[33m" // Yellow
		case LogLevelError:
			colorCode = "\033[31m" // Red
		case LogLevelCritical:
			colorCode = "\033[35m" // Magenta
		}
		fmt.Fprintf(h.writer, "%s%s\033[0m\n", colorCode, message)
	} else {
		fmt.Fprintln(h.writer, message)
	}
}

// NewLogger creates a new logger
func NewLogger(level LogLevel, output io.Writer) *Logger {
	return &Logger{
		level:     level,
		output:    output,
		formatter: NewDefaultFormatter(true, true, true),
		handlers:  make([]LogHandler, 0),
		context:   make(map[string]interface{}),
	}
}

// SetLevel sets the log level
func (l *Logger) SetLevel(level LogLevel) {
	l.level = level
}

// SetFormatter sets the log formatter
func (l *Logger) SetFormatter(formatter LogFormatter) {
	l.formatter = formatter
}

// AddHandler adds a log handler
func (l *Logger) AddHandler(handler LogHandler) {
	l.handlers = append(l.handlers, handler)
}

// SetContext sets context information
func (l *Logger) SetContext(key string, value interface{}) {
	l.context[key] = value
}

// EnableCallerInfo enables caller information in log messages
func (l *Logger) EnableCallerInfo(enable bool) {
	l.callerInfo = enable
}

// log logs a message at the specified level
func (l *Logger) log(level LogLevel, message string, args ...interface{}) {
	if level < l.level {
		return
	}
	
	// Format message
	formattedMessage := fmt.Sprintf(message, args...)
	
	// Get caller information
	var caller string
	if l.callerInfo {
		if pc, file, line, ok := runtime.Caller(2); ok {
			fn := runtime.FuncForPC(pc)
			caller = fmt.Sprintf("%s:%d %s", filepath.Base(file), line, fn.Name())
		}
	}
	
	// Format log message
	logMessage := l.formatter.Format(level, formattedMessage, l.context, caller)
	
	// Write to output
	if l.output != nil {
		fmt.Fprintln(l.output, logMessage)
	}
	
	// Handle with handlers
	for _, handler := range l.handlers {
		handler.Handle(level, formattedMessage, l.context)
	}
}

// Debug logs a debug message
func (l *Logger) Debug(message string, args ...interface{}) {
	l.log(LogLevelDebug, message, args...)
}

// Info logs an info message
func (l *Logger) Info(message string, args ...interface{}) {
	l.log(LogLevelInfo, message, args...)
}

// Warning logs a warning message
func (l *Logger) Warning(message string, args ...interface{}) {
	l.log(LogLevelWarning, message, args...)
}

// Error logs an error message
func (l *Logger) Error(message string, args ...interface{}) {
	l.log(LogLevelError, message, args...)
}

// Critical logs a critical message
func (l *Logger) Critical(message string, args ...interface{}) {
	l.log(LogLevelCritical, message, args...)
}

// LogManager provides centralized logging management
type LogManager struct {
	loggers map[string]*Logger
	config  *LogConfig
}

// LogConfig holds logging configuration
type LogConfig struct {
	DefaultLevel     LogLevel
	DefaultOutput    string
	DefaultFormatter string
	Handlers         []string
	Rotation         LogRotation
	Color            bool
}

// NewLogManager creates a new log manager
func NewLogManager() *LogManager {
	return &LogManager{
		loggers: make(map[string]*Logger),
		config: &LogConfig{
			DefaultLevel:     LogLevelInfo,
			DefaultOutput:    "console",
			DefaultFormatter: "default",
			Handlers:         []string{"console"},
			Rotation: LogRotation{
				MaxSize:    10 * 1024 * 1024, // 10MB
				MaxAge:     7 * 24 * time.Hour, // 7 days
				MaxBackups: 5,
			},
			Color: true,
		},
	}
}

// GetLogger gets or creates a logger
func (lm *LogManager) GetLogger(name string) *Logger {
	if logger, exists := lm.loggers[name]; exists {
		return logger
	}
	
	// Create new logger
	logger := NewLogger(lm.config.DefaultLevel, os.Stdout)
	
	// Add handlers based on configuration
	for _, handlerName := range lm.config.Handlers {
		switch handlerName {
		case "console":
			logger.AddHandler(NewConsoleHandler(lm.config.Color))
		case "file":
			if fileHandler, err := NewFileHandler("logs/"+name+".log", lm.config.Rotation); err == nil {
				logger.AddHandler(fileHandler)
			}
		}
	}
	
	lm.loggers[name] = logger
	return logger
}

// SetConfig sets the logging configuration
func (lm *LogManager) SetConfig(config *LogConfig) {
	lm.config = config
}

// Close closes all loggers
func (lm *LogManager) Close() {
    for _, logger := range lm.loggers {
        if file, ok := logger.output.(*os.File); ok {
            // 不要关闭标准输出或标准错误，否则会影响后续测试/程序输出
            if file == os.Stdout || file == os.Stderr {
                continue
            }
            _ = file.Close()
        }
    }
}

// PerformanceLogger provides performance logging functionality
type PerformanceLogger struct {
	logger *Logger
	start  time.Time
}

// NewPerformanceLogger creates a new performance logger
func NewPerformanceLogger(logger *Logger) *PerformanceLogger {
	return &PerformanceLogger{
		logger: logger,
		start:  time.Now(),
	}
}

// Start starts performance measurement
func (pl *PerformanceLogger) Start(operation string) {
	pl.start = time.Now()
	pl.logger.Debug("开始操作: %s", operation)
}

// End ends performance measurement
func (pl *PerformanceLogger) End(operation string) {
	duration := time.Since(pl.start)
	pl.logger.Info("操作完成: %s, 耗时: %v", operation, duration)
}

// LogOperation logs an operation with timing
func (pl *PerformanceLogger) LogOperation(operation string, fn func() error) error {
	pl.Start(operation)
	defer pl.End(operation)
	return fn()
}

// StructuredLogger provides structured logging functionality
type StructuredLogger struct {
	logger *Logger
	fields map[string]interface{}
}

// NewStructuredLogger creates a new structured logger
func NewStructuredLogger(logger *Logger) *StructuredLogger {
	return &StructuredLogger{
		logger: logger,
		fields: make(map[string]interface{}),
	}
}

// WithField adds a field to the structured logger
func (sl *StructuredLogger) WithField(key string, value interface{}) *StructuredLogger {
	sl.fields[key] = value
	return sl
}

// WithFields adds multiple fields to the structured logger
func (sl *StructuredLogger) WithFields(fields map[string]interface{}) *StructuredLogger {
	for key, value := range fields {
		sl.fields[key] = value
	}
	return sl
}

// Debug logs a debug message with fields
func (sl *StructuredLogger) Debug(message string) {
	sl.logger.SetContext("fields", sl.fields)
	sl.logger.Debug(message)
}

// Info logs an info message with fields
func (sl *StructuredLogger) Info(message string) {
	sl.logger.SetContext("fields", sl.fields)
	sl.logger.Info(message)
}

// Warning logs a warning message with fields
func (sl *StructuredLogger) Warning(message string) {
	sl.logger.SetContext("fields", sl.fields)
	sl.logger.Warning(message)
}

// Error logs an error message with fields
func (sl *StructuredLogger) Error(message string) {
	sl.logger.SetContext("fields", sl.fields)
	sl.logger.Error(message)
}

// Critical logs a critical message with fields
func (sl *StructuredLogger) Critical(message string) {
	sl.logger.SetContext("fields", sl.fields)
	sl.logger.Critical(message)
} 