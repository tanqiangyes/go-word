package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewDocumentProtection(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
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
	protection := wordprocessingml.NewDocumentProtection()
	
	settings := wordprocessingml.ProtectionSettings{
		Password:           "test123",
		EditRestrictions:   "readOnly",
		FormatRestrictions: "none",
		Enforcement:        true,
	}
	
	err := protection.EnableProtection(settings)
	if err != nil {
		t.Fatalf("Failed to enable protection: %v", err)
	}
	
	// 验证保护设置
	if protection.Settings.Password != "test123" {
		t.Errorf("Expected password 'test123', got '%s'", protection.Settings.Password)
	}
	
	if protection.Settings.EditRestrictions != "readOnly" {
		t.Errorf("Expected edit restrictions 'readOnly', got '%s'", protection.Settings.EditRestrictions)
	}
	
	if !protection.Settings.Enforcement {
		t.Error("Expected enforcement to be enabled")
	}
}

func TestDisableProtection(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	// 先启用保护
	settings := wordprocessingml.ProtectionSettings{
		Password:           "test123",
		EditRestrictions:   "readOnly",
		FormatRestrictions: "none",
		Enforcement:        true,
	}
	protection.EnableProtection(settings)
	
	// 禁用保护
	err := protection.DisableProtection("test123")
	if err != nil {
		t.Fatalf("Failed to disable protection: %v", err)
	}
	
	// 验证保护已禁用
	if protection.Settings.Enforcement {
		t.Error("Expected enforcement to be disabled")
	}
}

func TestAddUserPermission(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	userPermission := wordprocessingml.UserPermission{
		UserID:    "user1",
		Username:  "Test User",
		Email:     "test@example.com",
		Role:      "editor",
		Permissions: []string{"read", "write", "comment"},
		ExpiryDate: "2024-12-31",
	}
	
	err := protection.AddUserPermission(userPermission)
	if err != nil {
		t.Fatalf("Failed to add user permission: %v", err)
	}
	
	// 验证用户权限是否被添加
	permissions := protection.Permissions.UserPermissions
	if len(permissions) != 1 {
		t.Errorf("Expected 1 user permission, got %d", len(permissions))
	}
	
	if permissions[0].UserID != "user1" {
		t.Errorf("Expected user ID 'user1', got '%s'", permissions[0].UserID)
	}
}

func TestCheckPermission(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	// 添加用户权限
	userPermission := wordprocessingml.UserPermission{
		UserID:      "user1",
		Username:    "Test User",
		Role:        "editor",
		Permissions: []string{"read", "write", "comment"},
	}
	protection.AddUserPermission(userPermission)
	
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
	protection := wordprocessingml.NewDocumentProtection()
	
	watermark := wordprocessingml.Watermark{
		Text:      "CONFIDENTIAL",
		Type:      "text",
		Font:      "Arial",
		FontSize:  48,
		Color:     "FF0000",
		Opacity:   0.3,
		Rotation:  45,
		Position:  "center",
	}
	
	err := protection.AddWatermark(watermark)
	if err != nil {
		t.Fatalf("Failed to add watermark: %v", err)
	}
	
	// 验证水印是否被添加
	if protection.Watermark.Watermarks == nil || len(protection.Watermark.Watermarks) == 0 {
		t.Error("Expected watermark to be added")
	}
	
	addedWatermark := protection.Watermark.Watermarks[0]
	if addedWatermark.Text != "CONFIDENTIAL" {
		t.Errorf("Expected watermark text 'CONFIDENTIAL', got '%s'", addedWatermark.Text)
	}
	
	if addedWatermark.Type != "text" {
		t.Errorf("Expected watermark type 'text', got '%s'", addedWatermark.Type)
	}
}

func TestGetProtectionSummary(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	// 启用保护
	settings := wordprocessingml.ProtectionSettings{
		Password:           "test123",
		EditRestrictions:   "readOnly",
		FormatRestrictions: "none",
		Enforcement:        true,
	}
	protection.EnableProtection(settings)
	
	// 添加用户权限
	userPermission := wordprocessingml.UserPermission{
		UserID:      "user1",
		Username:    "Test User",
		Role:        "editor",
		Permissions: []string{"read", "write"},
	}
	protection.AddUserPermission(userPermission)
	
	// 添加水印
	watermark := wordprocessingml.Watermark{
		Text: "CONFIDENTIAL",
		Type: "text",
	}
	protection.AddWatermark(watermark)
	
	summary := protection.GetProtectionSummary()
	
	if summary == "" {
		t.Error("Expected non-empty protection summary")
	}
	
	// 检查摘要是否包含预期的保护信息
	expectedInfo := []string{"Protection Enabled", "User Permissions", "Watermarks", "Edit Restrictions"}
	for _, expected := range expectedInfo {
		if !contains(summary, expected) {
			t.Errorf("Expected summary to contain '%s'", expected)
		}
	}
}

func TestEncryptionSettings(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	// 设置加密
	encryptionSettings := wordprocessingml.EncryptionSettings{
		Algorithm:    "AES-256",
		KeySize:      256,
		Password:     "encrypt123",
		SaltLength:   16,
		SpinCount:    100000,
		HashAlgorithm: "SHA-256",
	}
	
	err := protection.Encryption.SetEncryptionSettings(encryptionSettings)
	if err != nil {
		t.Fatalf("Failed to set encryption settings: %v", err)
	}
	
	// 验证加密设置
	if protection.Encryption.Settings.Algorithm != "AES-256" {
		t.Errorf("Expected algorithm 'AES-256', got '%s'", protection.Encryption.Settings.Algorithm)
	}
	
	if protection.Encryption.Settings.KeySize != 256 {
		t.Errorf("Expected key size 256, got %d", protection.Encryption.Settings.KeySize)
	}
}

func TestDigitalSignatureSettings(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	// 设置数字签名
	signatureSettings := wordprocessingml.SignatureSettings{
		CertificatePath: "/path/to/certificate.pfx",
		CertificatePassword: "cert123",
		SignatureAlgorithm: "SHA-256",
		IncludeTimestamp: true,
		TimestampURL: "http://timestamp.digicert.com",
	}
	
	err := protection.DigitalSignature.SetSignatureSettings(signatureSettings)
	if err != nil {
		t.Fatalf("Failed to set signature settings: %v", err)
	}
	
	// 验证签名设置
	if protection.DigitalSignature.Settings.CertificatePath != "/path/to/certificate.pfx" {
		t.Errorf("Expected certificate path '/path/to/certificate.pfx', got '%s'", protection.DigitalSignature.Settings.CertificatePath)
	}
	
	if protection.DigitalSignature.Settings.SignatureAlgorithm != "SHA-256" {
		t.Errorf("Expected signature algorithm 'SHA-256', got '%s'", protection.DigitalSignature.Settings.SignatureAlgorithm)
	}
}

func TestPasswordHashing(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	password := "test123"
	hashedPassword := protection.hashPassword(password)
	
	if hashedPassword == "" {
		t.Error("Expected non-empty hashed password")
	}
	
	if hashedPassword == password {
		t.Error("Expected password to be hashed, not plain text")
	}
	
	// 测试相同密码产生相同的哈希
	hashedPassword2 := protection.hashPassword(password)
	if hashedPassword != hashedPassword2 {
		t.Error("Expected same password to produce same hash")
	}
}

func TestGroupPermission(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	groupPermission := wordprocessingml.GroupPermission{
		GroupID:     "group1",
		GroupName:   "Test Group",
		Role:        "reviewer",
		Permissions: []string{"read", "comment"},
		ExpiryDate:  "2024-12-31",
	}
	
	err := protection.Permissions.AddGroupPermission(groupPermission)
	if err != nil {
		t.Fatalf("Failed to add group permission: %v", err)
	}
	
	// 验证组权限是否被添加
	permissions := protection.Permissions.GroupPermissions
	if len(permissions) != 1 {
		t.Errorf("Expected 1 group permission, got %d", len(permissions))
	}
	
	if permissions[0].GroupID != "group1" {
		t.Errorf("Expected group ID 'group1', got '%s'", permissions[0].GroupID)
	}
}

func TestRolePermission(t *testing.T) {
	protection := wordprocessingml.NewDocumentProtection()
	
	rolePermission := wordprocessingml.RolePermission{
		RoleID:      "role1",
		RoleName:    "Editor Role",
		Description: "Can edit and comment",
		Permissions: []string{"read", "write", "comment"},
		Priority:    1,
	}
	
	err := protection.Permissions.AddRolePermission(rolePermission)
	if err != nil {
		t.Fatalf("Failed to add role permission: %v", err)
	}
	
	// 验证角色权限是否被添加
	permissions := protection.Permissions.RolePermissions
	if len(permissions) != 1 {
		t.Errorf("Expected 1 role permission, got %d", len(permissions))
	}
	
	if permissions[0].RoleID != "role1" {
		t.Errorf("Expected role ID 'role1', got '%s'", permissions[0].RoleID)
	}
}

// 辅助函数
