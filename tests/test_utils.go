package tests

import (
	"strings"
	"testing"
)

// contains checks if a string contains a substring
func contains(s, substr string) bool {
	return strings.Contains(s, substr)
}

// containsSubstring checks if a string contains a substring (case-insensitive)
func containsSubstring(s, substr string) bool {
	return strings.Contains(strings.ToLower(s), strings.ToLower(substr))
}

// assertContains asserts that a string contains a substring
func assertContains(t *testing.T, s, substr string) {
	if !contains(s, substr) {
		t.Errorf("Expected string to contain '%s', but got: %s", substr, s)
	}
}

// assertNotContains asserts that a string does not contain a substring
func assertNotContains(t *testing.T, s, substr string) {
	if contains(s, substr) {
		t.Errorf("Expected string to not contain '%s', but got: %s", substr, s)
	}
}

// assertEqual asserts that two values are equal
func assertEqual(t *testing.T, expected, actual interface{}) {
	if expected != actual {
		t.Errorf("Expected %v, but got %v", expected, actual)
	}
}

// assertNil asserts that a value is nil
func assertNil(t *testing.T, value interface{}) {
	if value != nil {
		t.Errorf("Expected nil, but got %v", value)
	}
}

// assertNotNil asserts that a value is not nil
func assertNotNil(t *testing.T, value interface{}) {
	if value == nil {
		t.Errorf("Expected not nil, but got nil")
	}
} 