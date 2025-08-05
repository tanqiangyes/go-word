// Package utils provides utility functions for the go-word library
package utils

import (
	"fmt"
	"strings"
)

// Error types for different categories of errors
const (
	ErrTypeParse     = "parse"
	ErrTypeIO        = "io"
	ErrTypeValidation = "validation"
	ErrTypeNotFound  = "not_found"
	ErrTypeCorrupt   = "corrupt"
)

// WordError represents a structured error with additional context
type WordError struct {
	Type    string
	Message string
	Cause   error
	Context map[string]interface{}
}

// Error implements the error interface
func (e *WordError) Error() string {
	if e.Cause != nil {
		return fmt.Sprintf("[%s] %s: %v", e.Type, e.Message, e.Cause)
	}
	return fmt.Sprintf("[%s] %s", e.Type, e.Message)
}

// Unwrap returns the underlying error
func (e *WordError) Unwrap() error {
	return e.Cause
}

// NewError creates a new WordError
func NewError(errType, message string, cause error) *WordError {
	return &WordError{
		Type:    errType,
		Message: message,
		Cause:   cause,
		Context: make(map[string]interface{}),
	}
}

// NewParseError creates a parse error
func NewParseError(message string, cause error) *WordError {
	return NewError(ErrTypeParse, message, cause)
}

// NewIOError creates an I/O error
func NewIOError(message string, cause error) *WordError {
	return NewError(ErrTypeIO, message, cause)
}

// NewValidationError creates a validation error
func NewValidationError(message string, cause error) *WordError {
	return NewError(ErrTypeValidation, message, cause)
}

// NewNotFoundError creates a not found error
func NewNotFoundError(message string, cause error) *WordError {
	return NewError(ErrTypeNotFound, message, cause)
}

// NewCorruptError creates a corrupt file error
func NewCorruptError(message string, cause error) *WordError {
	return NewError(ErrTypeCorrupt, message, cause)
}

// IsParseError checks if an error is a parse error
func IsParseError(err error) bool {
	return isErrorType(err, ErrTypeParse)
}

// IsIOError checks if an error is an I/O error
func IsIOError(err error) bool {
	return isErrorType(err, ErrTypeIO)
}

// IsValidationError checks if an error is a validation error
func IsValidationError(err error) bool {
	return isErrorType(err, ErrTypeValidation)
}

// IsNotFoundError checks if an error is a not found error
func IsNotFoundError(err error) bool {
	return isErrorType(err, ErrTypeNotFound)
}

// IsCorruptError checks if an error is a corrupt file error
func IsCorruptError(err error) bool {
	return isErrorType(err, ErrTypeCorrupt)
}

// isErrorType checks if an error is of a specific type
func isErrorType(err error, errType string) bool {
	if err == nil {
		return false
	}
	
	if wordErr, ok := err.(*WordError); ok {
		return wordErr.Type == errType
	}
	
	// Check error message for type indicator
	errMsg := err.Error()
	return strings.Contains(errMsg, fmt.Sprintf("[%s]", errType))
} 