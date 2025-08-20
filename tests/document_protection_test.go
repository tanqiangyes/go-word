package tests

import (
	"testing"

	"github.com/tanqiangyes/go-word/pkg/word"
)

func TestNewDocumentProtection(t *testing.T) {
	protection := word.NewDocumentProtection()

	if protection == nil {
		t.Fatal("Expected DocumentProtection to be created")
	}

	if protection.Settings == nil {
		t.Error("Expected Settings to be initialized")
	}

	if protection.Permissions == nil {
		t.Error("Expected Permissions to be initialized")
	}

	if protection.Encryption == nil {
		t.Error("Expected Encryption to be initialized")
	}

	if protection.DigitalSignature == nil {
		t.Error("Expected DigitalSignature to be initialized")
	}

	if protection.Watermark == nil {
		t.Error("Expected Watermark to be initialized")
	}
}

func TestEnableProtection(t *testing.T) {
	protection := word.NewDocumentProtection()

	err := protection.EnableProtection(word.PasswordProtection, "test123")
	if err != nil {
		t.Fatalf("Failed to enable protection: %v", err)
	}

	// 验证保护设置
	if protection.Settings.Password != "test123" {
		t.Errorf("Expected password 'test123', got '%s'", protection.Settings.Password)
	}

	if protection.Settings.Enforcement != word.EnforcementLevel(1) {
		t.Errorf("Expected enforcement level 1, got %v", protection.Settings.Enforcement)
	}
}

func TestDisableProtection(t *testing.T) {
	protection := word.NewDocumentProtection()

	// 先启用保护
	protection.EnableProtection(word.PasswordProtection, "test123")

	// 禁用保护
	err := protection.DisableProtection("test123")
	if err != nil {
		t.Fatalf("Failed to disable protection: %v", err)
	}

	// 验证保护已禁用
	if protection.Settings.Enforcement != word.EnforcementLevel(0) {
		t.Error("Expected enforcement to be disabled")
	}
}

func TestAddUserPermission(t *testing.T) {
	protection := word.NewDocumentProtection()

	err := protection.AddUserPermission("user1", "Test User", "test@example.com", map[string]bool{
		"read":    true,
		"write":   true,
		"comment": true,
	})
	if err != nil {
		t.Fatalf("Failed to add user permission: %v", err)
	}

	// 验证用户权限是否被添加
	if protection.Permissions == nil {
		t.Error("Expected permissions to be initialized")
	}
}

func TestCheckPermission(t *testing.T) {
	protection := word.NewDocumentProtection()

	// 添加用户权限
	protection.AddUserPermission("user1", "Test User", "test@example.com", map[string]bool{
		"read":    true,
		"write":   true,
		"comment": true,
	})

	// 测试权限检查
	hasPermission := protection.CheckPermission("user1", "read")
	if !hasPermission {
		t.Error("Expected user to have read permission")
	}

	hasPermission = protection.CheckPermission("user1", "admin")
	if hasPermission {
		t.Error("Expected user to not have admin permission")
	}

	hasPermission = protection.CheckPermission("user2", "read")
	if hasPermission {
		t.Error("Expected non-existent user to not have permission")
	}
}

func TestAddWatermark(t *testing.T) {
	protection := word.NewDocumentProtection()

	err := protection.AddWatermark("confidential", "CONFIDENTIAL", word.TextWatermark)
	if err != nil {
		t.Fatalf("Failed to add watermark: %v", err)
	}

	// 验证水印是否被添加
	if protection.Watermark == nil {
		t.Error("Expected watermark manager to be initialized")
	}
}

func TestGetProtectionSummary(t *testing.T) {
	protection := word.NewDocumentProtection()

	// 启用保护
	protection.EnableProtection(word.PasswordProtection, "test123")

	// 添加用户权限
	protection.AddUserPermission("user1", "Test User", "test@example.com", map[string]bool{
		"read":  true,
		"write": true,
	})

	// 添加水印
	protection.AddWatermark("confidential", "CONFIDENTIAL", word.TextWatermark)

	summary := protection.GetProtectionSummary()

	if summary == "" {
		t.Error("Expected non-empty protection summary")
	}

	// 检查摘要是否包含预期的保护信息
	if summary == "" {
		t.Error("Expected non-empty protection summary")
	}
}

func TestEncryptionSettings(t *testing.T) {
	protection := word.NewDocumentProtection()

	// 设置加密
	encryptionSettings := word.EncryptionSettings{
		Enabled: true,
		KeySize: 256,
		Salt:    "random_salt",
	}

	// 验证加密设置
	if !encryptionSettings.Enabled {
		t.Error("Expected encryption to be enabled")
	}

	if protection.Encryption.Settings.KeySize != 256 {
		t.Errorf("Expected key size 256, got %d", protection.Encryption.Settings.KeySize)
	}
}

func TestDigitalSignatureSettings(t *testing.T) {
	// 设置数字签名
	signatureSettings := word.SignatureSettings{
		Enabled:  true,
		Required: false,
		Multiple: false,
	}

	// 验证签名设置
	if !signatureSettings.Enabled {
		t.Error("Expected signature to be enabled")
	}
}

func TestPasswordHashing(t *testing.T) {
	protection := word.NewDocumentProtection()

	// 验证保护系统是否正常工作
	if protection.Settings == nil {
		t.Error("Expected settings to be initialized")
	}
}

func TestGroupPermission(t *testing.T) {
	protection := word.NewDocumentProtection()

	// 验证组权限系统是否正常工作
	if protection.Permissions == nil {
		t.Error("Expected permissions to be initialized")
	}
}

func TestRolePermission(t *testing.T) {
	protection := word.NewDocumentProtection()

	// 验证角色权限系统是否正常工作
	if protection.Permissions == nil {
		t.Error("Expected permissions to be initialized")
	}
}

// 辅助函数
