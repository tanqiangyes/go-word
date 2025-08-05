// Package utils provides utility functions for the go-word library.
// This package includes error handling utilities and common helper functions.
package utils

import (
	"fmt"
	"strings"
)

// Error types for different categories of errors

// DocumentError represents an error related to document operations
type DocumentError struct {
	Message string
	Cause   error
	Details map[string]interface{}
}

// ParseError represents an error during XML parsing
type ParseError struct {
	Message string
	Cause   error
	Line    int
	Column  int
}

// ValidationError represents an error during document validation
type ValidationError struct {
	Message string
	Cause   error
	Field   string
}

// IOError represents an error during file operations
type IOError struct {
	Message   string
	Cause     error
	Path      string
	Operation string
}

// FormatError represents an error related to document formatting
type FormatError struct {
	Message string
	Cause   error
	Element string
}

// NewError creates a new error with a message
func NewError(message string) error {
	return fmt.Errorf(message)
}

// WrapError wraps an existing error with additional context
func WrapError(err error, message string) error {
	if err == nil {
		return nil
	}
	return fmt.Errorf("%s: %w", message, err)
}

// NewDocumentError creates a new document-related error
func NewDocumentError(message string, cause error) *DocumentError {
	return &DocumentError{
		Message: message,
		Cause:   cause,
		Details: make(map[string]interface{}),
	}
}

// NewDocumentErrorWithDetails creates a new document error with additional details
func NewDocumentErrorWithDetails(message string, cause error, details map[string]interface{}) *DocumentError {
	return &DocumentError{
		Message: message,
		Cause:   cause,
		Details: details,
	}
}

// Error returns the error message
func (e *DocumentError) Error() string {
	if e.Cause != nil {
		return fmt.Sprintf("document error: %s: %v", e.Message, e.Cause)
	}
	return fmt.Sprintf("document error: %s", e.Message)
}

// Unwrap returns the underlying error
func (e *DocumentError) Unwrap() error {
	return e.Cause
}

// AddDetail adds a detail to the error
func (e *DocumentError) AddDetail(key string, value interface{}) {
	if e.Details == nil {
		e.Details = make(map[string]interface{})
	}
	e.Details[key] = value
}

// GetDetail returns a detail value
func (e *DocumentError) GetDetail(key string) interface{} {
	if e.Details == nil {
		return nil
	}
	return e.Details[key]
}

// NewParseError creates a new parsing error
func NewParseError(message string, cause error, line, column int) *ParseError {
	return &ParseError{
		Message: message,
		Cause:   cause,
		Line:    line,
		Column:  column,
	}
}

// Error returns the error message
func (e *ParseError) Error() string {
	if e.Line > 0 && e.Column > 0 {
		if e.Cause != nil {
			return fmt.Sprintf("parse error at line %d, column %d: %s: %v", 
				e.Line, e.Column, e.Message, e.Cause)
		}
		return fmt.Sprintf("parse error at line %d, column %d: %s", 
			e.Line, e.Column, e.Message)
	}
	if e.Cause != nil {
		return fmt.Sprintf("parse error: %s: %v", e.Message, e.Cause)
	}
	return fmt.Sprintf("parse error: %s", e.Message)
}

// Unwrap returns the underlying error
func (e *ParseError) Unwrap() error {
	return e.Cause
}

// NewValidationError creates a new validation error
func NewValidationError(message string, cause error, field string) *ValidationError {
	return &ValidationError{
		Message: message,
		Cause:   cause,
		Field:   field,
	}
}

// Error returns the error message
func (e *ValidationError) Error() string {
	if e.Field != "" {
		if e.Cause != nil {
			return fmt.Sprintf("validation error in field '%s': %s: %v", 
				e.Field, e.Message, e.Cause)
		}
		return fmt.Sprintf("validation error in field '%s': %s", e.Field, e.Message)
	}
	if e.Cause != nil {
		return fmt.Sprintf("validation error: %s: %v", e.Message, e.Cause)
	}
	return fmt.Sprintf("validation error: %s", e.Message)
}

// Unwrap returns the underlying error
func (e *ValidationError) Unwrap() error {
	return e.Cause
}

// NewIOError creates a new I/O error
func NewIOError(message string, cause error, path, operation string) *IOError {
	return &IOError{
		Message:   message,
		Cause:     cause,
		Path:      path,
		Operation: operation,
	}
}

// Error returns the error message
func (e *IOError) Error() string {
	if e.Path != "" && e.Operation != "" {
		if e.Cause != nil {
			return fmt.Sprintf("I/O error during %s on '%s': %s: %v", 
				e.Operation, e.Path, e.Message, e.Cause)
		}
		return fmt.Sprintf("I/O error during %s on '%s': %s", 
			e.Operation, e.Path, e.Message)
	}
	if e.Cause != nil {
		return fmt.Sprintf("I/O error: %s: %v", e.Message, e.Cause)
	}
	return fmt.Sprintf("I/O error: %s", e.Message)
}

// Unwrap returns the underlying error
func (e *IOError) Unwrap() error {
	return e.Cause
}

// NewFormatError creates a new formatting error
func NewFormatError(message string, cause error, element string) *FormatError {
	return &FormatError{
		Message: message,
		Cause:   cause,
		Element: element,
	}
}

// Error returns the error message
func (e *FormatError) Error() string {
	if e.Element != "" {
		if e.Cause != nil {
			return fmt.Sprintf("format error in element '%s': %s: %v", 
				e.Element, e.Message, e.Cause)
		}
		return fmt.Sprintf("format error in element '%s': %s", e.Element, e.Message)
	}
	if e.Cause != nil {
		return fmt.Sprintf("format error: %s: %v", e.Message, e.Cause)
	}
	return fmt.Sprintf("format error: %s", e.Message)
}

// Unwrap returns the underlying error
func (e *FormatError) Unwrap() error {
	return e.Cause
}

// Error utilities

// IsDocumentError checks if an error is a DocumentError
func IsDocumentError(err error) bool {
	_, ok := err.(*DocumentError)
	return ok
}

// IsParseError checks if an error is a ParseError
func IsParseError(err error) bool {
	_, ok := err.(*ParseError)
	return ok
}

// IsValidationError checks if an error is a ValidationError
func IsValidationError(err error) bool {
	_, ok := err.(*ValidationError)
	return ok
}

// IsIOError checks if an error is an IOError
func IsIOError(err error) bool {
	_, ok := err.(*IOError)
	return ok
}

// IsFormatError checks if an error is a FormatError
func IsFormatError(err error) bool {
	_, ok := err.(*FormatError)
	return ok
}

// GetErrorType returns the type of error as a string
func GetErrorType(err error) string {
	switch {
	case IsDocumentError(err):
		return "DocumentError"
	case IsParseError(err):
		return "ParseError"
	case IsValidationError(err):
		return "ValidationError"
	case IsIOError(err):
		return "IOError"
	case IsFormatError(err):
		return "FormatError"
	default:
		return "UnknownError"
	}
}

// GetUserFriendlyMessage returns a user-friendly error message
func GetUserFriendlyMessage(err error) string {
	if err == nil {
		return ""
	}

	switch e := err.(type) {
	case *DocumentError:
		return getUserFriendlyDocumentError(e)
	case *ParseError:
		return getUserFriendlyParseError(e)
	case *ValidationError:
		return getUserFriendlyValidationError(e)
	case *IOError:
		return getUserFriendlyIOError(e)
	case *FormatError:
		return getUserFriendlyFormatError(e)
	default:
		return fmt.Sprintf("发生错误: %v", err)
	}
}

// getUserFriendlyDocumentError returns a user-friendly document error message
func getUserFriendlyDocumentError(err *DocumentError) string {
	message := "文档处理错误"
	
	if strings.Contains(err.Message, "not found") {
		message = "文档文件未找到"
	} else if strings.Contains(err.Message, "corrupted") {
		message = "文档文件已损坏"
	} else if strings.Contains(err.Message, "unsupported") {
		message = "不支持的文档格式"
	} else if strings.Contains(err.Message, "permission") {
		message = "没有权限访问文档"
	}
	
	if err.Cause != nil {
		return fmt.Sprintf("%s: %v", message, err.Cause)
	}
	return message
}

// getUserFriendlyParseError returns a user-friendly parse error message
func getUserFriendlyParseError(err *ParseError) string {
	message := "文档解析错误"
	
	if err.Line > 0 && err.Column > 0 {
		message = fmt.Sprintf("文档解析错误 (第%d行, 第%d列)", err.Line, err.Column)
	}
	
	if err.Cause != nil {
		return fmt.Sprintf("%s: %v", message, err.Cause)
	}
	return message
}

// getUserFriendlyValidationError returns a user-friendly validation error message
func getUserFriendlyValidationError(err *ValidationError) string {
	message := "文档验证错误"
	
	if err.Field != "" {
		message = fmt.Sprintf("文档验证错误 (字段: %s)", err.Field)
	}
	
	if err.Cause != nil {
		return fmt.Sprintf("%s: %v", message, err.Cause)
	}
	return message
}

// getUserFriendlyIOError returns a user-friendly I/O error message
func getUserFriendlyIOError(err *IOError) string {
	message := "文件操作错误"
	
	if err.Operation != "" {
		switch err.Operation {
		case "read":
			message = "文件读取错误"
		case "write":
			message = "文件写入错误"
		case "open":
			message = "文件打开错误"
		case "close":
			message = "文件关闭错误"
		}
	}
	
	if err.Path != "" {
		message = fmt.Sprintf("%s (文件: %s)", message, err.Path)
	}
	
	if err.Cause != nil {
		return fmt.Sprintf("%s: %v", message, err.Cause)
	}
	return message
}

// getUserFriendlyFormatError returns a user-friendly format error message
func getUserFriendlyFormatError(err *FormatError) string {
	message := "文档格式错误"
	
	if err.Element != "" {
		message = fmt.Sprintf("文档格式错误 (元素: %s)", err.Element)
	}
	
	if err.Cause != nil {
		return fmt.Sprintf("%s: %v", message, err.Cause)
	}
	return message
}

// Error context utilities

// AddErrorContext adds context information to an error
func AddErrorContext(err error, context map[string]interface{}) error {
	if err == nil {
		return nil
	}
	
	switch e := err.(type) {
	case *DocumentError:
		for k, v := range context {
			e.AddDetail(k, v)
		}
		return e
	default:
		// For other error types, wrap with additional context
		contextStr := ""
		for k, v := range context {
			if contextStr != "" {
				contextStr += ", "
			}
			contextStr += fmt.Sprintf("%s=%v", k, v)
		}
		return WrapError(err, fmt.Sprintf("context: %s", contextStr))
	}
}

// GetErrorContext extracts context information from an error
func GetErrorContext(err error) map[string]interface{} {
	if err == nil {
		return nil
	}
	
	switch e := err.(type) {
	case *DocumentError:
		if e.Details == nil {
			return make(map[string]interface{})
		}
		return e.Details
	default:
		return make(map[string]interface{})
	}
} 