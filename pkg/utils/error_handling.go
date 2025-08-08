// Package utils provides utility functions and error handling
package utils

import (
	"fmt"
	"runtime"
	"strings"
)

// ErrorCode represents a specific error code
type ErrorCode string

const (
	// Document errors
	ErrDocumentNotFound     ErrorCode = "DOCUMENT_NOT_FOUND"
	ErrDocumentCorrupted    ErrorCode = "DOCUMENT_CORRUPTED"
	ErrDocumentInvalid      ErrorCode = "DOCUMENT_INVALID"
	ErrDocumentProtected    ErrorCode = "DOCUMENT_PROTECTED"
	ErrDocumentReadOnly     ErrorCode = "DOCUMENT_READ_ONLY"
	
	// File system errors
	ErrFileNotFound        ErrorCode = "FILE_NOT_FOUND"
	ErrFilePermission      ErrorCode = "FILE_PERMISSION"
	ErrFileCorrupted       ErrorCode = "FILE_CORRUPTED"
	ErrFileTooLarge        ErrorCode = "FILE_TOO_LARGE"
	
	// XML parsing errors
	ErrXMLParseFailed      ErrorCode = "XML_PARSE_FAILED"
	ErrXMLInvalid          ErrorCode = "XML_INVALID"
	ErrXMLUnsupported      ErrorCode = "XML_UNSUPPORTED"
	
	// OPC container errors
	ErrOPCInvalid          ErrorCode = "OPC_INVALID"
	ErrOPCCorrupted        ErrorCode = "OPC_CORRUPTED"
	ErrOPCMissingPart      ErrorCode = "OPC_MISSING_PART"
	
	// Content errors
	ErrContentInvalid      ErrorCode = "CONTENT_INVALID"
	ErrContentTooLarge     ErrorCode = "CONTENT_TOO_LARGE"
	ErrContentUnsupported  ErrorCode = "CONTENT_UNSUPPORTED"
	
	// Style errors
	ErrStyleNotFound       ErrorCode = "STYLE_NOT_FOUND"
	ErrStyleInvalid        ErrorCode = "STYLE_INVALID"
	ErrStyleConflict       ErrorCode = "STYLE_CONFLICT"
	
	// Protection errors
	ErrProtectionFailed    ErrorCode = "PROTECTION_FAILED"
	ErrPasswordIncorrect   ErrorCode = "PASSWORD_INCORRECT"
	ErrPermissionDenied    ErrorCode = "PERMISSION_DENIED"
	
	// Validation errors
	ErrValidationFailed    ErrorCode = "VALIDATION_FAILED"
	ErrValidationRule      ErrorCode = "VALIDATION_RULE"
	
	// Format errors
	ErrFormatUnsupported   ErrorCode = "FORMAT_UNSUPPORTED"
	ErrFormatConversion    ErrorCode = "FORMAT_CONVERSION"
	
	// Memory errors
	ErrMemoryInsufficient  ErrorCode = "MEMORY_INSUFFICIENT"
	ErrMemoryAllocation    ErrorCode = "MEMORY_ALLOCATION"
	
	// System errors
	ErrSystemResource      ErrorCode = "SYSTEM_RESOURCE"
	ErrSystemTimeout       ErrorCode = "SYSTEM_TIMEOUT"
	ErrSystemUnavailable   ErrorCode = "SYSTEM_UNAVAILABLE"
	
	// Resource errors
	ErrResourceExhausted   ErrorCode = "RESOURCE_EXHAUSTED"
	ErrInvalidInput        ErrorCode = "INVALID_INPUT"
	ErrInvalidState        ErrorCode = "INVALID_STATE"
)

// ErrorSeverity represents the severity level of an error
type ErrorSeverity int

const (
	SeverityInfo ErrorSeverity = iota
	SeverityWarning
	SeverityError
	SeverityCritical
)

// StructuredDocumentError represents a structured document-related error
type StructuredDocumentError struct {
	Code       ErrorCode
	Message    string
	Severity   ErrorSeverity
	Context    map[string]interface{}
	InnerError error
	Stack      []string
}

// NewStructuredDocumentError creates a new structured document error
func NewStructuredDocumentError(code ErrorCode, message string) *StructuredDocumentError {
	return &StructuredDocumentError{
		Code:     code,
		Message:  message,
		Severity: SeverityError,
		Context:  make(map[string]interface{}),
		Stack:    getStackTrace(),
	}
}

// WithSeverity sets the error severity
func (e *StructuredDocumentError) WithSeverity(severity ErrorSeverity) *StructuredDocumentError {
	e.Severity = severity
	return e
}

// WithContext adds context information to the error
func (e *StructuredDocumentError) WithContext(key string, value interface{}) *StructuredDocumentError {
	e.Context[key] = value
	return e
}

// WithInnerError sets the inner error
func (e *StructuredDocumentError) WithInnerError(err error) *StructuredDocumentError {
	e.InnerError = err
	return e
}

// Error implements the error interface
func (e *StructuredDocumentError) Error() string {
	var parts []string
	
	// Add error code
	parts = append(parts, fmt.Sprintf("[%s]", e.Code))
	
	// Add message
	parts = append(parts, e.Message)
	
	// Add context if available
	if len(e.Context) > 0 {
		var contextParts []string
		for key, value := range e.Context {
			contextParts = append(contextParts, fmt.Sprintf("%s=%v", key, value))
		}
		parts = append(parts, fmt.Sprintf("(%s)", strings.Join(contextParts, ", ")))
	}
	
	// Add inner error if available
	if e.InnerError != nil {
		parts = append(parts, fmt.Sprintf("caused by: %v", e.InnerError))
	}
	
	return strings.Join(parts, " ")
}

// Unwrap returns the inner error
func (e *StructuredDocumentError) Unwrap() error {
	return e.InnerError
}

// IsErrorCode checks if an error has a specific error code
func IsErrorCode(err error, code ErrorCode) bool {
	if docErr, ok := err.(*StructuredDocumentError); ok {
		return docErr.Code == code
	}
	return false
}

// GetErrorCode returns the error code if available
func GetErrorCode(err error) ErrorCode {
	if docErr, ok := err.(*StructuredDocumentError); ok {
		return docErr.Code
	}
	return ""
}

// GetErrorSeverity returns the error severity if available
func GetErrorSeverity(err error) ErrorSeverity {
	if docErr, ok := err.(*StructuredDocumentError); ok {
		return docErr.Severity
	}
	return SeverityError
}

// ErrorContext provides context for error creation
type ErrorContext struct {
	DocumentPath string
	Operation    string
	LineNumber   int
	Function     string
	Parameters   map[string]interface{}
}

// NewErrorContext creates a new error context
func NewErrorContext() *ErrorContext {
	return &ErrorContext{
		Parameters: make(map[string]interface{}),
	}
}

// WithDocumentPath sets the document path
func (ctx *ErrorContext) WithDocumentPath(path string) *ErrorContext {
	ctx.DocumentPath = path
	return ctx
}

// WithOperation sets the operation being performed
func (ctx *ErrorContext) WithOperation(operation string) *ErrorContext {
	ctx.Operation = operation
	return ctx
}

// WithParameter adds a parameter to the context
func (ctx *ErrorContext) WithParameter(key string, value interface{}) *ErrorContext {
	ctx.Parameters[key] = value
	return ctx
}

// ErrorHandler provides centralized error handling
type ErrorHandler struct {
	handlers map[ErrorCode]ErrorHandlerFunc
	logger   ErrorLogger
}

// ErrorHandlerFunc is a function that handles a specific error
type ErrorHandlerFunc func(error, *ErrorContext) error

// ErrorLogger provides error logging functionality
type ErrorLogger interface {
	LogError(err error, context *ErrorContext)
	LogWarning(err error, context *ErrorContext)
	LogInfo(err error, context *ErrorContext)
}

// NewErrorHandler creates a new error handler
func NewErrorHandler() *ErrorHandler {
	return &ErrorHandler{
		handlers: make(map[ErrorCode]ErrorHandlerFunc),
	}
}

// RegisterHandler registers an error handler for a specific error code
func (h *ErrorHandler) RegisterHandler(code ErrorCode, handler ErrorHandlerFunc) {
	h.handlers[code] = handler
}

// SetLogger sets the error logger
func (h *ErrorHandler) SetLogger(logger ErrorLogger) {
	h.logger = logger
}

// HandleError handles an error using registered handlers
func (h *ErrorHandler) HandleError(err error, context *ErrorContext) error {
	if docErr, ok := err.(*StructuredDocumentError); ok {
		if handler, exists := h.handlers[docErr.Code]; exists {
			return handler(err, context)
		}
	}
	
	// Log error if logger is available
	if h.logger != nil {
		h.logger.LogError(err, context)
	}
	
	return err
}

// getStackTrace returns the current stack trace
func getStackTrace() []string {
	var stack []string
	for i := 1; i < 10; i++ {
		pc, file, line, ok := runtime.Caller(i)
		if !ok {
			break
		}
		fn := runtime.FuncForPC(pc)
		stack = append(stack, fmt.Sprintf("%s:%d %s", file, line, fn.Name()))
	}
	return stack
}

// ErrorRecovery provides error recovery mechanisms
type ErrorRecovery struct {
	handler *ErrorHandler
}

// NewErrorRecovery creates a new error recovery system
func NewErrorRecovery(handler *ErrorHandler) *ErrorRecovery {
	return &ErrorRecovery{
		handler: handler,
	}
}

// RecoverFromError attempts to recover from an error
func (r *ErrorRecovery) RecoverFromError(err error, context *ErrorContext) error {
	// Try to handle the error
	handledErr := r.handler.HandleError(err, context)
	
	// If the error is still a document error, try recovery strategies
	if docErr, ok := handledErr.(*StructuredDocumentError); ok {
		switch docErr.Code {
		case ErrDocumentCorrupted:
			return r.recoverFromCorruption(docErr, context)
		case ErrMemoryInsufficient:
			return r.recoverFromMemoryError(docErr, context)
		case ErrSystemTimeout:
			return r.recoverFromTimeout(docErr, context)
		}
	}
	
	return handledErr
}

// recoverFromCorruption attempts to recover from document corruption
func (r *ErrorRecovery) recoverFromCorruption(err *StructuredDocumentError, context *ErrorContext) error {
	// Try to extract valid content from corrupted document
	// This is a simplified implementation
	return NewStructuredDocumentError(ErrDocumentCorrupted, "无法从损坏的文档中恢复")
}

// recoverFromMemoryError attempts to recover from memory errors
func (r *ErrorRecovery) recoverFromMemoryError(err *StructuredDocumentError, context *ErrorContext) error {
	// Try to free memory and retry
	// This is a simplified implementation
	return NewStructuredDocumentError(ErrMemoryInsufficient, "内存不足，无法继续操作")
}

// recoverFromTimeout attempts to recover from timeout errors
func (r *ErrorRecovery) recoverFromTimeout(err *StructuredDocumentError, context *ErrorContext) error {
	// Try to extend timeout or retry
	// This is a simplified implementation
	return NewStructuredDocumentError(ErrSystemTimeout, "操作超时，请重试")
}

// ErrorMetrics provides error metrics collection
type ErrorMetrics struct {
	errors    map[ErrorCode]int
	severities map[ErrorSeverity]int
	total     int
}

// NewErrorMetrics creates new error metrics
func NewErrorMetrics() *ErrorMetrics {
	return &ErrorMetrics{
		errors:     make(map[ErrorCode]int),
		severities: make(map[ErrorSeverity]int),
	}
}

// RecordError records an error for metrics
func (m *ErrorMetrics) RecordError(err error) {
	m.total++
	
	if docErr, ok := err.(*StructuredDocumentError); ok {
		m.errors[docErr.Code]++
		m.severities[docErr.Severity]++
	}
}

// GetMetrics returns the current error metrics
func (m *ErrorMetrics) GetMetrics() map[string]interface{} {
	return map[string]interface{}{
		"total_errors": m.total,
		"error_codes":  m.errors,
		"severities":   m.severities,
	}
} 