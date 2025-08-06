// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"crypto/md5"
	"encoding/hex"
	"fmt"
	"strings"
	"time"
)

// DocumentProtection represents document protection functionality
type DocumentProtection struct {
	// 保护设置
	Settings *ProtectionSettings
	
	// 权限管理
	Permissions *PermissionManager
	
	// 加密功能
	Encryption *EncryptionManager
	
	// 数字签名
	DigitalSignature *DigitalSignatureManager
	
	// 水印功能
	Watermark *WatermarkManager
}

// ProtectionSettings represents document protection settings
type ProtectionSettings struct {
	// 基础设置
	Enabled     bool
	ProtectionType ProtectionType
	Password    string
	Hash        string
	
	// 保护选项
	EditRestrictions EditRestrictions
	FormatRestrictions FormatRestrictions
	
	// 其他设置
	Enforcement EnforcementLevel
	GracePeriod time.Duration
}

// ProtectionType defines the type of protection
type ProtectionType int

const (
	// NoProtection for no protection
	NoProtection ProtectionType = iota
	// ReadOnlyProtection for read-only protection
	ReadOnlyProtection
	// CommentsProtection for comments-only protection
	CommentsProtection
	// TrackChangesProtection for track changes protection
	TrackChangesProtection
	// FormsProtection for forms protection
	FormsProtection
	// PasswordProtection for password protection
	PasswordProtection
)

// EditRestrictions represents edit restrictions
type EditRestrictions struct {
	// 编辑限制
	AllowEditing bool
	AllowDeletion bool
	AllowInsertion bool
	AllowFormatting bool
	AllowPrinting bool
	
	// 特定限制
	RestrictToComments bool
	RestrictToTrackChanges bool
	RestrictToForms bool
	RestrictToReadOnly bool
}

// FormatRestrictions represents format restrictions
type FormatRestrictions struct {
	// 格式限制
	AllowFontFormatting bool
	AllowParagraphFormatting bool
	AllowSectionFormatting bool
	AllowTableFormatting bool
	AllowHeaderFooterFormatting bool
	
	// 其他限制
	AllowStyleFormatting bool
	AllowThemeFormatting bool
	AllowPageSetup bool
}

// EnforcementLevel defines the enforcement level
type EnforcementLevel int

const (
	// NoEnforcement for no enforcement
	NoEnforcement EnforcementLevel = iota
	// SoftEnforcement for soft enforcement
	SoftEnforcement
	// HardEnforcement for hard enforcement
	HardEnforcement
)

// PermissionManager manages document permissions
type PermissionManager struct {
	// 用户权限
	UserPermissions map[string]*UserPermission
	
	// 组权限
	GroupPermissions map[string]*GroupPermission
	
	// 角色权限
	RolePermissions map[string]*RolePermission
	
	// 权限策略
	Policies []PermissionPolicy
}

// UserPermission represents user permission
type UserPermission struct {
	// 基础信息
	UserID      string
	UserName    string
	Email       string
	
	// 权限设置
	CanRead     bool
	CanEdit     bool
	CanDelete   bool
	CanFormat   bool
	CanPrint    bool
	CanShare    bool
	
	// 时间限制
	ValidFrom   time.Time
	ValidUntil  time.Time
	
	// 其他属性
	Description string
	Priority    int
}

// GroupPermission represents group permission
type GroupPermission struct {
	// 基础信息
	GroupID     string
	GroupName   string
	Description string
	
	// 权限设置
	CanRead     bool
	CanEdit     bool
	CanDelete   bool
	CanFormat   bool
	CanPrint    bool
	CanShare    bool
	
	// 成员管理
	Members     []string
	InheritFrom string
}

// RolePermission represents role permission
type RolePermission struct {
	// 基础信息
	RoleID      string
	RoleName    string
	Description string
	
	// 权限设置
	CanRead     bool
	CanEdit     bool
	CanDelete   bool
	CanFormat   bool
	CanPrint    bool
	CanShare    bool
	CanAdmin    bool
	
	// 继承设置
	InheritFrom string
	Priority    int
}

// PermissionPolicy represents a permission policy
type PermissionPolicy struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 策略设置
	Type        PolicyType
	Conditions  []PolicyCondition
	Actions     []PolicyAction
	
	// 属性
	Enabled     bool
	Priority    int
}

// PolicyType defines the type of policy
type PolicyType int

const (
	// AllowPolicy for allow policies
	AllowPolicy PolicyType = iota
	// DenyPolicy for deny policies
	DenyPolicy
	// RequirePolicy for require policies
	RequirePolicy
)

// PolicyCondition represents a policy condition
type PolicyCondition struct {
	// 基础信息
	ID          string
	Type        ConditionType
	Value       string
	
	// 属性
	Operator    ConditionOperator
	Negated     bool
}

// ConditionType defines the type of condition
type ConditionType int

const (
	// UserCondition for user conditions
	UserCondition ConditionType = iota
	// TimeCondition for time conditions
	TimeCondition
	// LocationCondition for location conditions
	LocationCondition
	// DeviceCondition for device conditions
	DeviceCondition
)

// ConditionOperator defines the condition operator
type ConditionOperator int

const (
	// EqualsOperator for equals
	EqualsOperator ConditionOperator = iota
	// NotEqualsOperator for not equals
	NotEqualsOperator
	// ContainsOperator for contains
	ContainsOperator
	// GreaterThanOperator for greater than
	GreaterThanOperator
	// LessThanOperator for less than
	LessThanOperator
)

// PolicyAction represents a policy action
type PolicyAction struct {
	// 基础信息
	ID          string
	Type        ActionType
	Value       string
	
	// 属性
	Priority    int
	Enabled     bool
}

// ActionType defines the type of action
type ActionType int

const (
	// GrantAction for grant actions
	GrantAction ActionType = iota
	// DenyAction for deny actions
	DenyAction
	// LogAction for log actions
	LogAction
	// NotifyAction for notify actions
	NotifyAction
)

// EncryptionManager manages document encryption
type EncryptionManager struct {
	// 加密设置
	Settings *EncryptionSettings
	
	// 加密算法
	Algorithm EncryptionAlgorithm
	
	// 密钥管理
	KeyManager *KeyManager
	
	// 加密历史
	History []EncryptionRecord
}

// EncryptionSettings represents encryption settings
type EncryptionSettings struct {
	// 基础设置
	Enabled     bool
	Algorithm   EncryptionAlgorithm
	KeySize     int
	Salt        string
	
	// 加密选项
	EncryptContent bool
	EncryptMetadata bool
	EncryptHeaders bool
	
	// 其他设置
	Compression bool
	BackupKey   bool
}

// EncryptionAlgorithm defines the encryption algorithm
type EncryptionAlgorithm int

const (
	// AES128Algorithm for AES-128
	AES128Algorithm EncryptionAlgorithm = iota
	// AES256Algorithm for AES-256
	AES256Algorithm
	// RC4Algorithm for RC4
	RC4Algorithm
	// DESAlgorithm for DES
	DESAlgorithm
)

// KeyManager manages encryption keys
type KeyManager struct {
	// 密钥存储
	Keys map[string]*EncryptionKey
	
	// 主密钥
	MasterKey *EncryptionKey
	
	// 密钥策略
	Policies []KeyPolicy
}

// EncryptionKey represents an encryption key
type EncryptionKey struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 密钥数据
	Key         []byte
	IV          []byte
	Salt        []byte
	
	// 属性
	Algorithm   EncryptionAlgorithm
	KeySize     int
	Created     time.Time
	Expires     time.Time
	
	// 其他属性
	Active      bool
	Protected   bool
}

// KeyPolicy represents a key policy
type KeyPolicy struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 策略设置
	MinKeySize  int
	MaxKeySize  int
	Algorithm   EncryptionAlgorithm
	Rotation    time.Duration
	
	// 属性
	Enabled     bool
	Priority    int
}

// EncryptionRecord represents an encryption record
type EncryptionRecord struct {
	// 基础信息
	ID          string
	Timestamp   time.Time
	UserID      string
	
	// 操作详情
	Operation   EncryptionOperation
	Algorithm   EncryptionAlgorithm
	KeyID       string
	
	// 结果
	Success     bool
	Error       string
}

// EncryptionOperation defines the encryption operation
type EncryptionOperation int

const (
	// EncryptOperation for encrypt
	EncryptOperation EncryptionOperation = iota
	// DecryptOperation for decrypt
	DecryptOperation
	// KeyRotationOperation for key rotation
	KeyRotationOperation
	// KeyBackupOperation for key backup
	KeyBackupOperation
)

// DigitalSignatureManager manages digital signatures
type DigitalSignatureManager struct {
	// 签名设置
	Settings *SignatureSettings
	
	// 证书管理
	Certificates map[string]*Certificate
	
	// 签名历史
	Signatures []DigitalSignature
	
	// 验证设置
	Validation *SignatureValidationSettings
}

// SignatureSettings represents signature settings
type SignatureSettings struct {
	// 基础设置
	Enabled     bool
	Required    bool
	Multiple    bool
	
	// 签名选项
	SignContent bool
	SignMetadata bool
	SignHeaders bool
	
	// 其他设置
	Timestamp   bool
	CertificateValidation bool
}

// Certificate represents a digital certificate
type Certificate struct {
	// 基础信息
	ID          string
	Subject     string
	Issuer      string
	SerialNumber string
	
	// 证书数据
	PublicKey   []byte
	PrivateKey  []byte
	Certificate []byte
	
	// 属性
	ValidFrom   time.Time
	ValidUntil  time.Time
	KeySize     int
	Algorithm   string
	
	// 其他属性
	Active      bool
	Trusted     bool
}

// DigitalSignature represents a digital signature
type DigitalSignature struct {
	// 基础信息
	ID          string
	Timestamp   time.Time
	SignerID    string
	SignerName  string
	
	// 签名数据
	Signature   []byte
	Hash        string
	Algorithm   string
	
	// 属性
	Valid       bool
	Trusted     bool
	Revoked     bool
	
	// 其他属性
	Description string
	Location    string
}

// SignatureValidationSettings represents signature validation settings
type SignatureValidationSettings struct {
	// 验证设置
	CheckRevocation bool
	CheckExpiration bool
	CheckTrust      bool
	
	// 验证策略
	Policies []ValidationPolicy
	
	// 验证结果
	Results map[string]*ValidationResult
}

// ValidationPolicy represents a validation policy
type ValidationPolicy struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 策略设置
	Type        ValidationType
	Conditions  []ValidationCondition
	
	// 属性
	Enabled     bool
	Priority    int
}

// ValidationType defines the type of validation
type ValidationType int

const (
	// CertificateValidation for certificate validation
	CertificateValidation ValidationType = iota
	// SignatureValidation for signature validation
	SignatureValidation
	// TimestampValidation for timestamp validation
	TimestampValidation
)

// ValidationCondition represents a validation condition
type ValidationCondition struct {
	// 基础信息
	ID          string
	Type        ConditionType
	Value       string
	
	// 属性
	Operator    ConditionOperator
	Required    bool
}

// ValidationResult 使用 document_validator.go 中的定义

// WatermarkManager manages document watermarks
type WatermarkManager struct {
	// 水印设置
	Settings *WatermarkSettings
	
	// 水印列表
	Watermarks []Watermark
	
	// 水印模板
	Templates map[string]*WatermarkTemplate
}

// WatermarkSettings represents watermark settings
type WatermarkSettings struct {
	// 基础设置
	Enabled     bool
	DefaultTemplate string
	
	// 显示选项
	ShowOnPrint bool
	ShowOnScreen bool
	ShowOnExport bool
	
	// 其他设置
	Opacity     float64
	Rotation    float64
	Position    WatermarkPosition
}

// WatermarkPosition defines watermark position
type WatermarkPosition int

const (
	// CenterPosition for center
	CenterPosition WatermarkPosition = iota
	// TopLeftPosition for top left
	TopLeftPosition
	// TopRightPosition for top right
	TopRightPosition
	// BottomLeftPosition for bottom left
	BottomLeftPosition
	// BottomRightPosition for bottom right
	BottomRightPosition
)

// Watermark represents a watermark
type Watermark struct {
	// 基础信息
	ID          string
	Name        string
	Type        WatermarkType
	Text        string
	
	// 样式设置
	Font        string
	Size        int
	Color       string
	Opacity     float64
	Rotation    float64
	
	// 位置设置
	Position    WatermarkPosition
	X           float64
	Y           float64
	
	// 其他属性
	Active      bool
	Created     time.Time
}

// WatermarkType defines the type of watermark
type WatermarkType int

const (
	// TextWatermark for text watermarks
	TextWatermark WatermarkType = iota
	// ImageWatermark for image watermarks
	ImageWatermark
	// LogoWatermark for logo watermarks
	LogoWatermark
	// CustomWatermark for custom watermarks
	CustomWatermark
)

// WatermarkTemplate represents a watermark template
type WatermarkTemplate struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 模板设置
	Type        WatermarkType
	Text        string
	Image       []byte
	
	// 样式设置
	Font        string
	Size        int
	Color       string
	Opacity     float64
	Rotation    float64
	
	// 位置设置
	Position    WatermarkPosition
	X           float64
	Y           float64
	
	// 其他属性
	Active      bool
	Default     bool
}

// NewDocumentProtection creates new document protection
func NewDocumentProtection() *DocumentProtection {
	return &DocumentProtection{
		Settings: &ProtectionSettings{
			Enabled: false,
			ProtectionType: NoProtection,
			EditRestrictions: EditRestrictions{
				AllowEditing: true,
				AllowDeletion: true,
				AllowInsertion: true,
				AllowFormatting: true,
				AllowPrinting: true,
			},
			FormatRestrictions: FormatRestrictions{
				AllowFontFormatting: true,
				AllowParagraphFormatting: true,
				AllowSectionFormatting: true,
				AllowTableFormatting: true,
				AllowHeaderFooterFormatting: true,
				AllowStyleFormatting: true,
				AllowThemeFormatting: true,
				AllowPageSetup: true,
			},
			Enforcement: NoEnforcement,
		},
		Permissions: &PermissionManager{
			UserPermissions: make(map[string]*UserPermission),
			GroupPermissions: make(map[string]*GroupPermission),
			RolePermissions: make(map[string]*RolePermission),
			Policies: make([]PermissionPolicy, 0),
		},
		Encryption: &EncryptionManager{
			Settings: &EncryptionSettings{
				Enabled: false,
				Algorithm: AES256Algorithm,
				KeySize: 256,
				EncryptContent: true,
				EncryptMetadata: false,
				EncryptHeaders: false,
				Compression: true,
				BackupKey: false,
			},
			Algorithm: AES256Algorithm,
			KeyManager: &KeyManager{
				Keys: make(map[string]*EncryptionKey),
				Policies: make([]KeyPolicy, 0),
			},
			History: make([]EncryptionRecord, 0),
		},
		DigitalSignature: &DigitalSignatureManager{
			Settings: &SignatureSettings{
				Enabled: false,
				Required: false,
				Multiple: false,
				SignContent: true,
				SignMetadata: false,
				SignHeaders: false,
				Timestamp: true,
				CertificateValidation: true,
			},
			Certificates: make(map[string]*Certificate),
			Signatures: make([]DigitalSignature, 0),
			Validation: &SignatureValidationSettings{
				CheckRevocation: true,
				CheckExpiration: true,
				CheckTrust: true,
				Policies: make([]ValidationPolicy, 0),
				Results: make(map[string]*ValidationResult),
			},
		},
		Watermark: &WatermarkManager{
			Settings: &WatermarkSettings{
				Enabled: false,
				ShowOnPrint: true,
				ShowOnScreen: true,
				ShowOnExport: true,
				Opacity: 0.5,
				Rotation: 0.0,
				Position: CenterPosition,
			},
			Watermarks: make([]Watermark, 0),
			Templates: make(map[string]*WatermarkTemplate),
		},
	}
}

// EnableProtection enables document protection
func (dp *DocumentProtection) EnableProtection(protectionType ProtectionType, password string) error {
	if dp.Settings == nil {
		return fmt.Errorf("protection settings not initialized")
	}
	
	dp.Settings.Enabled = true
	dp.Settings.ProtectionType = protectionType
	dp.Settings.Enforcement = SoftEnforcement // 设置强制级别为软强制
	
	if password != "" {
		dp.Settings.Password = password
		dp.Settings.Hash = dp.hashPassword(password)
	}
	
	// 根据保护类型设置相应的限制
	switch protectionType {
	case ReadOnlyProtection:
		dp.Settings.EditRestrictions.AllowEditing = false
		dp.Settings.EditRestrictions.AllowDeletion = false
		dp.Settings.EditRestrictions.AllowInsertion = false
		dp.Settings.EditRestrictions.AllowFormatting = false
	case CommentsProtection:
		dp.Settings.EditRestrictions.RestrictToComments = true
		dp.Settings.EditRestrictions.AllowEditing = false
		dp.Settings.EditRestrictions.AllowDeletion = false
		dp.Settings.EditRestrictions.AllowInsertion = false
	case TrackChangesProtection:
		dp.Settings.EditRestrictions.RestrictToTrackChanges = true
		dp.Settings.EditRestrictions.AllowEditing = false
	case FormsProtection:
		dp.Settings.EditRestrictions.RestrictToForms = true
		dp.Settings.EditRestrictions.AllowEditing = false
		dp.Settings.EditRestrictions.AllowFormatting = false
	}
	
	return nil
}

// DisableProtection disables document protection
func (dp *DocumentProtection) DisableProtection(password string) error {
	if !dp.Settings.Enabled {
		return fmt.Errorf("document protection is not enabled")
	}
	
	if dp.Settings.Password != "" {
		if dp.Settings.Hash != dp.hashPassword(password) {
			return fmt.Errorf("incorrect password")
		}
	}
	
	dp.Settings.Enabled = false
	dp.Settings.ProtectionType = NoProtection
	dp.Settings.Password = ""
	dp.Settings.Hash = ""
	
	// 重置所有限制
	dp.Settings.EditRestrictions = EditRestrictions{
		AllowEditing: true,
		AllowDeletion: true,
		AllowInsertion: true,
		AllowFormatting: true,
		AllowPrinting: true,
	}
	
	dp.Settings.FormatRestrictions = FormatRestrictions{
		AllowFontFormatting: true,
		AllowParagraphFormatting: true,
		AllowSectionFormatting: true,
		AllowTableFormatting: true,
		AllowHeaderFooterFormatting: true,
		AllowStyleFormatting: true,
		AllowThemeFormatting: true,
		AllowPageSetup: true,
	}
	
	return nil
}

// CheckPermission checks if a user has permission for an action
func (dp *DocumentProtection) CheckPermission(userID, action string) bool {
	// 如果保护没有启用，仍然检查用户权限
	// 如果没有找到用户权限，返回 false
	
	// 检查用户权限
	if userPerm := dp.Permissions.UserPermissions[userID]; userPerm != nil {
		switch action {
		case "read":
			return userPerm.CanRead
		case "edit", "write":
			return userPerm.CanEdit
		case "delete":
			return userPerm.CanDelete
		case "format":
			return userPerm.CanFormat
		case "print":
			return userPerm.CanPrint
		case "share":
			return userPerm.CanShare
		case "comment":
			return userPerm.CanFormat // 评论权限映射到格式权限
		default:
			// 对于不存在的权限类型，返回 false
			return false
		}
	}
	
	// 检查组权限
	for _, groupPerm := range dp.Permissions.GroupPermissions {
		for _, member := range groupPerm.Members {
			if member == userID {
				switch action {
				case "read":
					return groupPerm.CanRead
				case "edit":
					return groupPerm.CanEdit
				case "delete":
					return groupPerm.CanDelete
				case "format":
					return groupPerm.CanFormat
				case "print":
					return groupPerm.CanPrint
				case "share":
					return groupPerm.CanShare
				}
			}
		}
	}
	
	// 如果没有找到用户权限或组权限，返回 false
	return false
}

// AddUserPermission adds user permission
func (dp *DocumentProtection) AddUserPermission(userID, userName, email string, permissions map[string]bool) error {
	if dp.Permissions == nil {
		return fmt.Errorf("permission manager not initialized")
	}
	
	userPerm := &UserPermission{
		UserID:    userID,
		UserName:  userName,
		Email:     email,
		CanRead:   permissions["read"],
		CanEdit:   permissions["edit"] || permissions["write"], // 支持 write 作为 edit 的别名
		CanDelete: permissions["delete"],
		CanFormat: permissions["format"] || permissions["comment"], // 支持 comment 作为 format 的别名
		CanPrint:  permissions["print"],
		CanShare:  permissions["share"],
		ValidFrom: time.Now(),
		ValidUntil: time.Now().AddDate(1, 0, 0), // 默认一年有效期
		Priority:  1,
	}
	
	dp.Permissions.UserPermissions[userID] = userPerm
	return nil
}

// AddWatermark adds a watermark
func (dp *DocumentProtection) AddWatermark(name, text string, watermarkType WatermarkType) error {
	if dp.Watermark == nil {
		return fmt.Errorf("watermark manager not initialized")
	}
	
	watermark := Watermark{
		ID:       fmt.Sprintf("watermark_%d", len(dp.Watermark.Watermarks)+1),
		Name:     name,
		Type:     watermarkType,
		Text:     text,
		Font:     "Arial",
		Size:     48,
		Color:    "808080",
		Opacity:  0.5,
		Rotation: -45.0,
		Position: CenterPosition,
		X:        0.0,
		Y:        0.0,
		Active:   true,
		Created:  time.Now(),
	}
	
	dp.Watermark.Watermarks = append(dp.Watermark.Watermarks, watermark)
	return nil
}

// GetProtectionSummary returns a summary of protection settings
func (dp *DocumentProtection) GetProtectionSummary() string {
	var summary strings.Builder
	summary.WriteString("文档保护摘要:\n")
	
	if dp.Settings != nil {
		summary.WriteString(fmt.Sprintf("保护状态: %v\n", dp.Settings.Enabled))
		summary.WriteString(fmt.Sprintf("保护类型: %v\n", dp.Settings.ProtectionType))
		summary.WriteString(fmt.Sprintf("密码保护: %v\n", dp.Settings.Password != ""))
		summary.WriteString(fmt.Sprintf("强制级别: %v\n", dp.Settings.Enforcement))
	}
	
	if dp.Permissions != nil {
		summary.WriteString(fmt.Sprintf("用户权限: %d\n", len(dp.Permissions.UserPermissions)))
		summary.WriteString(fmt.Sprintf("组权限: %d\n", len(dp.Permissions.GroupPermissions)))
		summary.WriteString(fmt.Sprintf("角色权限: %d\n", len(dp.Permissions.RolePermissions)))
		summary.WriteString(fmt.Sprintf("权限策略: %d\n", len(dp.Permissions.Policies)))
	}
	
	if dp.Encryption != nil {
		summary.WriteString(fmt.Sprintf("加密状态: %v\n", dp.Encryption.Settings.Enabled))
		summary.WriteString(fmt.Sprintf("加密算法: %v\n", dp.Encryption.Settings.Algorithm))
		summary.WriteString(fmt.Sprintf("密钥数量: %d\n", len(dp.Encryption.KeyManager.Keys)))
		summary.WriteString(fmt.Sprintf("加密记录: %d\n", len(dp.Encryption.History)))
	}
	
	if dp.DigitalSignature != nil {
		summary.WriteString(fmt.Sprintf("数字签名: %v\n", dp.DigitalSignature.Settings.Enabled))
		summary.WriteString(fmt.Sprintf("证书数量: %d\n", len(dp.DigitalSignature.Certificates)))
		summary.WriteString(fmt.Sprintf("签名数量: %d\n", len(dp.DigitalSignature.Signatures)))
	}
	
	if dp.Watermark != nil {
		summary.WriteString(fmt.Sprintf("水印状态: %v\n", dp.Watermark.Settings.Enabled))
		summary.WriteString(fmt.Sprintf("水印数量: %d\n", len(dp.Watermark.Watermarks)))
		summary.WriteString(fmt.Sprintf("模板数量: %d\n", len(dp.Watermark.Templates)))
	}
	
	return summary.String()
}

// hashPassword hashes a password
func (dp *DocumentProtection) hashPassword(password string) string {
	hash := md5.Sum([]byte(password))
	return hex.EncodeToString(hash[:])
} 