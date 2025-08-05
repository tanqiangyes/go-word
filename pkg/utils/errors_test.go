package utils

import (
	"errors"
	"testing"
)

func TestNewError(t *testing.T) {
	err := NewError("test error")
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if err.Error() != "test error" {
		t.Errorf("Expected error message 'test error', got '%s'", err.Error())
	}
}

func TestWrapError(t *testing.T) {
	cause := errors.New("underlying error")
	err := WrapError(cause, "wrapped error")
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	errorStr := err.Error()
	if errorStr == "" {
		t.Error("Expected error string to not be empty")
	}
}

func TestNewDocumentError(t *testing.T) {
	cause := errors.New("underlying error")
	err := NewDocumentError("document error", cause)
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if !IsDocumentError(err) {
		t.Error("Expected IsDocumentError to return true")
	}
	
	// 测试错误消息
	errorStr := err.Error()
	if errorStr == "" {
		t.Error("Expected error string to not be empty")
	}
}

func TestNewParseError(t *testing.T) {
	cause := errors.New("parse failed")
	err := NewParseError("failed to parse document", cause, 10, 25)
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if !IsParseError(err) {
		t.Error("Expected IsParseError to return true")
	}
	
	// 测试错误消息
	errorStr := err.Error()
	if errorStr == "" {
		t.Error("Expected error string to not be empty")
	}
}

func TestNewIOError(t *testing.T) {
	cause := errors.New("file not found")
	err := NewIOError("failed to read file", cause, "/path/to/file", "read")
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if !IsIOError(err) {
		t.Error("Expected IsIOError to return true")
	}
	
	// 测试错误消息
	errorStr := err.Error()
	if errorStr == "" {
		t.Error("Expected error string to not be empty")
	}
}

func TestNewValidationError(t *testing.T) {
	cause := errors.New("invalid format")
	err := NewValidationError("document validation failed", cause, "content")
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if !IsValidationError(err) {
		t.Error("Expected IsValidationError to return true")
	}
	
	// 测试错误消息
	errorStr := err.Error()
	if errorStr == "" {
		t.Error("Expected error string to not be empty")
	}
}

func TestNewFormatError(t *testing.T) {
	cause := errors.New("format failed")
	err := NewFormatError("format error", cause, "paragraph")
	
	if err == nil {
		t.Fatal("Expected error to be created")
	}
	
	if !IsFormatError(err) {
		t.Error("Expected IsFormatError to return true")
	}
	
	// 测试错误消息
	errorStr := err.Error()
	if errorStr == "" {
		t.Error("Expected error string to not be empty")
	}
}

func TestDocumentErrorAddDetail(t *testing.T) {
	err := NewDocumentError("test error", nil)
	
	// 添加详情
	err.AddDetail("filename", "test.docx")
	err.AddDetail("size", 1024)
	
	// 获取详情
	filename := err.GetDetail("filename")
	if filename != "test.docx" {
		t.Errorf("Expected filename 'test.docx', got '%v'", filename)
	}
	
	size := err.GetDetail("size")
	if size != 1024 {
		t.Errorf("Expected size 1024, got '%v'", size)
	}
}

func TestErrorUnwrap(t *testing.T) {
	cause := errors.New("underlying error")
	err := NewDocumentError("test error", cause)
	
	unwrapped := err.Unwrap()
	if unwrapped != cause {
		t.Error("Expected unwrapped error to match cause")
	}
}

func TestIsParseError(t *testing.T) {
	cause := errors.New("parse failed")
	err := NewParseError("parse error", cause, 1, 1)
	
	if !IsParseError(err) {
		t.Error("Expected IsParseError to return true")
	}
	
	if IsIOError(err) {
		t.Error("Expected IsIOError to return false")
	}
}

func TestIsIOError(t *testing.T) {
	cause := errors.New("file not found")
	err := NewIOError("io error", cause, "/path", "read")
	
	if !IsIOError(err) {
		t.Error("Expected IsIOError to return true")
	}
	
	if IsParseError(err) {
		t.Error("Expected IsParseError to return false")
	}
}

func TestIsValidationError(t *testing.T) {
	cause := errors.New("invalid format")
	err := NewValidationError("validation error", cause, "field")
	
	if !IsValidationError(err) {
		t.Error("Expected IsValidationError to return true")
	}
}

func TestIsFormatError(t *testing.T) {
	cause := errors.New("format failed")
	err := NewFormatError("format error", cause, "element")
	
	if !IsFormatError(err) {
		t.Error("Expected IsFormatError to return true")
	}
}

func TestGetErrorType(t *testing.T) {
	// 测试文档错误
	docErr := NewDocumentError("test", nil)
	if GetErrorType(docErr) != "DocumentError" {
		t.Errorf("Expected error type 'DocumentError', got '%s'", GetErrorType(docErr))
	}
	
	// 测试解析错误
	parseErr := NewParseError("test", nil, 1, 1)
	if GetErrorType(parseErr) != "ParseError" {
		t.Errorf("Expected error type 'ParseError', got '%s'", GetErrorType(parseErr))
	}
	
	// 测试未知错误
	unknownErr := errors.New("unknown error")
	if GetErrorType(unknownErr) != "UnknownError" {
		t.Errorf("Expected error type 'UnknownError', got '%s'", GetErrorType(unknownErr))
	}
}

func TestGetUserFriendlyMessage(t *testing.T) {
	// 测试文档错误
	docErr := NewDocumentError("document not found", nil)
	message := GetUserFriendlyMessage(docErr)
	if message == "" {
		t.Error("Expected user-friendly message to not be empty")
	}
	
	// 测试解析错误
	parseErr := NewParseError("invalid XML", nil, 10, 25)
	message = GetUserFriendlyMessage(parseErr)
	if message == "" {
		t.Error("Expected user-friendly message to not be empty")
	}
	
	// 测试I/O错误
	ioErr := NewIOError("permission denied", nil, "/path", "read")
	message = GetUserFriendlyMessage(ioErr)
	if message == "" {
		t.Error("Expected user-friendly message to not be empty")
	}
}

func TestAddErrorContext(t *testing.T) {
	err := NewDocumentError("test error", nil)
	
	context := map[string]interface{}{
		"filename": "test.docx",
		"size":     1024,
	}
	
	contextErr := AddErrorContext(err, context)
	if contextErr == nil {
		t.Error("Expected error with context to be created")
	}
	
	// 验证上下文被添加
	if docErr, ok := contextErr.(*DocumentError); ok {
		filename := docErr.GetDetail("filename")
		if filename != "test.docx" {
			t.Errorf("Expected filename 'test.docx', got '%v'", filename)
		}
	}
}

func TestGetErrorContext(t *testing.T) {
	err := NewDocumentError("test error", nil)
	
	// 添加一些详情
	err.AddDetail("filename", "test.docx")
	err.AddDetail("size", 1024)
	
	context := GetErrorContext(err)
	if context == nil {
		t.Error("Expected context to not be nil")
	}
	
	if context["filename"] != "test.docx" {
		t.Errorf("Expected filename 'test.docx', got '%v'", context["filename"])
	}
	
	if context["size"] != 1024 {
		t.Errorf("Expected size 1024, got '%v'", context["size"])
	}
}