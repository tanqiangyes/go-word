package utils

import (
	"errors"
	"testing"
)

func TestNewError(t *testing.T) {
	cause := errors.New("underlying error")
	err := NewError("test", "test error", cause)
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if err.Type != "test" {
		t.Errorf("Expected type 'test', got '%s'", err.Type)
	}
	
	if err.Message != "test error" {
		t.Errorf("Expected message 'test error', got '%s'", err.Message)
	}
	
	if err.Cause != cause {
		t.Error("Expected cause to match")
	}
}

func TestNewParseError(t *testing.T) {
	cause := errors.New("parse failed")
	err := NewParseError("failed to parse document", cause)
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if err.Type != ErrTypeParse {
		t.Errorf("Expected type '%s', got '%s'", ErrTypeParse, err.Type)
	}
	
	if err.Message != "failed to parse document" {
		t.Errorf("Expected message 'failed to parse document', got '%s'", err.Message)
	}
}

func TestNewIOError(t *testing.T) {
	cause := errors.New("file not found")
	err := NewIOError("failed to read file", cause)
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if err.Type != ErrTypeIO {
		t.Errorf("Expected type '%s', got '%s'", ErrTypeIO, err.Type)
	}
	
	if err.Message != "failed to read file" {
		t.Errorf("Expected message 'failed to read file', got '%s'", err.Message)
	}
}

func TestNewValidationError(t *testing.T) {
	cause := errors.New("invalid format")
	err := NewValidationError("document validation failed", cause)
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if err.Type != ErrTypeValidation {
		t.Errorf("Expected type '%s', got '%s'", ErrTypeValidation, err.Type)
	}
	
	if err.Message != "document validation failed" {
		t.Errorf("Expected message 'document validation failed', got '%s'", err.Message)
	}
}

func TestNewNotFoundError(t *testing.T) {
	cause := errors.New("file not found")
	err := NewNotFoundError("document not found", cause)
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if err.Type != ErrTypeNotFound {
		t.Errorf("Expected type '%s', got '%s'", ErrTypeNotFound, err.Type)
	}
	
	if err.Message != "document not found" {
		t.Errorf("Expected message 'document not found', got '%s'", err.Message)
	}
}

func TestNewCorruptError(t *testing.T) {
	cause := errors.New("corrupt data")
	err := NewCorruptError("document is corrupt", cause)
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if err.Type != ErrTypeCorrupt {
		t.Errorf("Expected type '%s', got '%s'", ErrTypeCorrupt, err.Type)
	}
	
	if err.Message != "document is corrupt" {
		t.Errorf("Expected message 'document is corrupt', got '%s'", err.Message)
	}
}

func TestErrorString(t *testing.T) {
	cause := errors.New("underlying error")
	err := NewError("test", "test error", cause)
	
	errorStr := err.Error()
	if errorStr == "" {
		t.Error("Expected error string to not be empty")
	}
	
	// 验证错误字符串包含类型和消息
	if err.Type != "test" {
		t.Error("Expected error type to be 'test'")
	}
}

func TestErrorUnwrap(t *testing.T) {
	cause := errors.New("underlying error")
	err := NewError("test", "test error", cause)
	
	unwrapped := err.Unwrap()
	if unwrapped != cause {
		t.Error("Expected unwrapped error to match cause")
	}
}

func TestIsParseError(t *testing.T) {
	cause := errors.New("parse failed")
	err := NewParseError("parse error", cause)
	
	if !IsParseError(err) {
		t.Error("Expected IsParseError to return true")
	}
	
	if IsIOError(err) {
		t.Error("Expected IsIOError to return false")
	}
}

func TestIsIOError(t *testing.T) {
	cause := errors.New("file not found")
	err := NewIOError("io error", cause)
	
	if !IsIOError(err) {
		t.Error("Expected IsIOError to return true")
	}
	
	if IsParseError(err) {
		t.Error("Expected IsParseError to return false")
	}
}

func TestIsValidationError(t *testing.T) {
	cause := errors.New("invalid format")
	err := NewValidationError("validation error", cause)
	
	if !IsValidationError(err) {
		t.Error("Expected IsValidationError to return true")
	}
}

func TestIsNotFoundError(t *testing.T) {
	cause := errors.New("not found")
	err := NewNotFoundError("not found error", cause)
	
	if !IsNotFoundError(err) {
		t.Error("Expected IsNotFoundError to return true")
	}
}

func TestIsCorruptError(t *testing.T) {
	cause := errors.New("corrupt")
	err := NewCorruptError("corrupt error", cause)
	
	if !IsCorruptError(err) {
		t.Error("Expected IsCorruptError to return true")
	}
}

func TestErrorWithContext(t *testing.T) {
	err := NewError("test", "test error", nil)
	
	// 验证Context字段存在
	if err.Context == nil {
		t.Error("Expected Context to be initialized")
	}
} 