// Package utils provides utility functions and configuration management
package utils

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// ConfigManager provides configuration management functionality
type ConfigManager struct {
	Config     *LibraryConfig
	ConfigPath string
	Overrides  map[string]interface{}
}

// LibraryConfig holds the library configuration
type LibraryConfig struct {
	// General settings
	Language        string `json:"language"`
	DefaultEncoding string `json:"default_encoding"`
	LogLevel        string `json:"log_level"`
	
	// Performance settings
	MaxFileSize     int64  `json:"max_file_size"`
	MemoryLimit     int64  `json:"memory_limit"`
	Timeout         int    `json:"timeout"`
	BufferSize      int    `json:"buffer_size"`
	
	// Document settings
	DefaultFont     string `json:"default_font"`
	DefaultFontSize int    `json:"default_font_size"`
	DefaultStyle    string `json:"default_style"`
	
	// Format settings
	SupportedFormats []string `json:"supported_formats"`
	AutoFormat       bool     `json:"auto_format"`
	
	// Security settings
	EnableProtection bool   `json:"enable_protection"`
	DefaultPassword  string `json:"default_password"`
	
	// Validation settings
	EnableValidation bool     `json:"enable_validation"`
	ValidationRules  []string `json:"validation_rules"`
	AutoFix          bool     `json:"auto_fix"`
	
	// Error handling settings
	ErrorRecovery    bool `json:"error_recovery"`
	ErrorLogging     bool `json:"error_logging"`
	ErrorMetrics     bool `json:"error_metrics"`
	
	// Advanced settings
	EnableCaching    bool `json:"enable_caching"`
	EnableCompression bool `json:"enable_compression"`
	EnableEncryption  bool `json:"enable_encryption"`
	
	// Compatibility settings
	WPSCompatibility bool `json:"wps_compatibility"`
	OfficeCompatibility bool `json:"office_compatibility"`
	
	// Debug settings
	DebugMode        bool `json:"debug_mode"`
	VerboseLogging   bool `json:"verbose_logging"`
	PerformanceProfiling bool `json:"performance_profiling"`
}

// NewConfigManager creates a new configuration manager
func NewConfigManager() *ConfigManager {
	return &ConfigManager{
		Config: &LibraryConfig{
			Language:        "zh-CN",
			DefaultEncoding: "UTF-8",
			LogLevel:        "info",
			MaxFileSize:     100 * 1024 * 1024, // 100MB
			MemoryLimit:     512 * 1024 * 1024, // 512MB
			Timeout:         30,
			BufferSize:      8192,
			DefaultFont:     "Microsoft YaHei",
			DefaultFontSize: 12,
			DefaultStyle:    "Normal",
			SupportedFormats: []string{".docx", ".doc", ".rtf"},
			AutoFormat:       true,
			EnableProtection: false,
			EnableValidation: true,
			ValidationRules:  []string{"basic", "format", "content"},
			AutoFix:          false,
			ErrorRecovery:    true,
			ErrorLogging:     true,
			ErrorMetrics:     true,
			EnableCaching:    true,
			EnableCompression: true,
			EnableEncryption:  false,
			WPSCompatibility: true,
			OfficeCompatibility: true,
			DebugMode:        false,
			VerboseLogging:   false,
			PerformanceProfiling: false,
		},
		Overrides: make(map[string]interface{}),
	}
}

// LoadConfig loads configuration from a file
func (cm *ConfigManager) LoadConfig(configPath string) error {
	cm.ConfigPath = configPath
	
	// Check if config file exists
	if _, err := os.Stat(configPath); os.IsNotExist(err) {
		// Create default config file
		return cm.SaveConfig()
	}
	
	// Read config file
	data, err := os.ReadFile(configPath)
	if err != nil {
		return fmt.Errorf("failed to read config file: %w", err)
	}
	
	// Parse JSON
	if err := json.Unmarshal(data, cm.Config); err != nil {
		return fmt.Errorf("failed to parse config file: %w", err)
	}
	
	return nil
}

// SaveConfig saves configuration to a file
func (cm *ConfigManager) SaveConfig() error {
	if cm.ConfigPath == "" {
		cm.ConfigPath = "go-word-config.json"
	}
	
	// Create directory if it doesn't exist
	dir := filepath.Dir(cm.ConfigPath)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return fmt.Errorf("failed to create config directory: %w", err)
	}
	
	// Marshal to JSON
	data, err := json.MarshalIndent(cm.Config, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to marshal config: %w", err)
	}
	
	// Write to file
	if err := os.WriteFile(cm.ConfigPath, data, 0644); err != nil {
		return fmt.Errorf("failed to write config file: %w", err)
	}
	
	return nil
}

// GetConfig returns the current configuration
func (cm *ConfigManager) GetConfig() *LibraryConfig {
	return cm.Config
}

// SetConfig sets a configuration value
func (cm *ConfigManager) SetConfig(key string, value interface{}) {
	cm.Overrides[key] = value
}

// GetString gets a string configuration value
func (cm *ConfigManager) GetString(key string) string {
	// Check overrides first
	if override, exists := cm.Overrides[key]; exists {
		if str, ok := override.(string); ok {
			return str
		}
	}
	
	// Check config struct
	switch key {
	case "language":
		return cm.Config.Language
	case "default_encoding":
		return cm.Config.DefaultEncoding
	case "log_level":
		return cm.Config.LogLevel
	case "default_font":
		return cm.Config.DefaultFont
	case "default_style":
		return cm.Config.DefaultStyle
	case "default_password":
		return cm.Config.DefaultPassword
	}
	
	return ""
}

// GetInt gets an integer configuration value
func (cm *ConfigManager) GetInt(key string) int {
	// Check overrides first
	if override, exists := cm.Overrides[key]; exists {
		if i, ok := override.(int); ok {
			return i
		}
		if str, ok := override.(string); ok {
			if i, err := strconv.Atoi(str); err == nil {
				return i
			}
		}
	}
	
	// Check config struct
	switch key {
	case "timeout":
		return cm.Config.Timeout
	case "buffer_size":
		return cm.Config.BufferSize
	case "default_font_size":
		return cm.Config.DefaultFontSize
	}
	
	return 0
}

// GetInt64 gets an int64 configuration value
func (cm *ConfigManager) GetInt64(key string) int64 {
	// Check overrides first
	if override, exists := cm.Overrides[key]; exists {
		if i, ok := override.(int64); ok {
			return i
		}
		if str, ok := override.(string); ok {
			if i, err := strconv.ParseInt(str, 10, 64); err == nil {
				return i
			}
		}
	}
	
	// Check config struct
	switch key {
	case "max_file_size":
		return cm.Config.MaxFileSize
	case "memory_limit":
		return cm.Config.MemoryLimit
	}
	
	return 0
}

// GetBool gets a boolean configuration value
func (cm *ConfigManager) GetBool(key string) bool {
	// Check overrides first
	if override, exists := cm.Overrides[key]; exists {
		if b, ok := override.(bool); ok {
			return b
		}
		if str, ok := override.(string); ok {
			return strings.ToLower(str) == "true"
		}
	}
	
	// Check config struct
	switch key {
	case "auto_format":
		return cm.Config.AutoFormat
	case "enable_protection":
		return cm.Config.EnableProtection
	case "enable_validation":
		return cm.Config.EnableValidation
	case "auto_fix":
		return cm.Config.AutoFix
	case "error_recovery":
		return cm.Config.ErrorRecovery
	case "error_logging":
		return cm.Config.ErrorLogging
	case "error_metrics":
		return cm.Config.ErrorMetrics
	case "enable_caching":
		return cm.Config.EnableCaching
	case "enable_compression":
		return cm.Config.EnableCompression
	case "enable_encryption":
		return cm.Config.EnableEncryption
	case "wps_compatibility":
		return cm.Config.WPSCompatibility
	case "office_compatibility":
		return cm.Config.OfficeCompatibility
	case "debug_mode":
		return cm.Config.DebugMode
	case "verbose_logging":
		return cm.Config.VerboseLogging
	case "performance_profiling":
		return cm.Config.PerformanceProfiling
	}
	
	return false
}

// GetStringSlice gets a string slice configuration value
func (cm *ConfigManager) GetStringSlice(key string) []string {
	// Check overrides first
	if override, exists := cm.Overrides[key]; exists {
		if slice, ok := override.([]string); ok {
			return slice
		}
	}
	
	// Check config struct
	switch key {
	case "supported_formats":
		return cm.Config.SupportedFormats
	case "validation_rules":
		return cm.Config.ValidationRules
	}
	
	return []string{}
}

// EnvironmentConfig loads configuration from environment variables
type EnvironmentConfig struct {
	prefix string
}

// NewEnvironmentConfig creates a new environment configuration loader
func NewEnvironmentConfig(prefix string) *EnvironmentConfig {
	return &EnvironmentConfig{
		prefix: prefix,
	}
}

// LoadFromEnvironment loads configuration from environment variables
func (ec *EnvironmentConfig) LoadFromEnvironment(cm *ConfigManager) {
	// Load string values
	if val := os.Getenv(ec.prefix + "LANGUAGE"); val != "" {
		cm.SetConfig("language", val)
	}
	if val := os.Getenv(ec.prefix + "DEFAULT_ENCODING"); val != "" {
		cm.SetConfig("default_encoding", val)
	}
	if val := os.Getenv(ec.prefix + "LOG_LEVEL"); val != "" {
		cm.SetConfig("log_level", val)
	}
	if val := os.Getenv(ec.prefix + "DEFAULT_FONT"); val != "" {
		cm.SetConfig("default_font", val)
	}
	if val := os.Getenv(ec.prefix + "DEFAULT_STYLE"); val != "" {
		cm.SetConfig("default_style", val)
	}
	
	// Load integer values
	if val := os.Getenv(ec.prefix + "TIMEOUT"); val != "" {
		if i, err := strconv.Atoi(val); err == nil {
			cm.SetConfig("timeout", i)
		}
	}
	if val := os.Getenv(ec.prefix + "BUFFER_SIZE"); val != "" {
		if i, err := strconv.Atoi(val); err == nil {
			cm.SetConfig("buffer_size", i)
		}
	}
	if val := os.Getenv(ec.prefix + "DEFAULT_FONT_SIZE"); val != "" {
		if i, err := strconv.Atoi(val); err == nil {
			cm.SetConfig("default_font_size", i)
		}
	}
	
	// Load int64 values
	if val := os.Getenv(ec.prefix + "MAX_FILE_SIZE"); val != "" {
		if i, err := strconv.ParseInt(val, 10, 64); err == nil {
			cm.SetConfig("max_file_size", i)
		}
	}
	if val := os.Getenv(ec.prefix + "MEMORY_LIMIT"); val != "" {
		if i, err := strconv.ParseInt(val, 10, 64); err == nil {
			cm.SetConfig("memory_limit", i)
		}
	}
	
	// Load boolean values
	if val := os.Getenv(ec.prefix + "AUTO_FORMAT"); val != "" {
		cm.SetConfig("auto_format", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "ENABLE_PROTECTION"); val != "" {
		cm.SetConfig("enable_protection", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "ENABLE_VALIDATION"); val != "" {
		cm.SetConfig("enable_validation", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "AUTO_FIX"); val != "" {
		cm.SetConfig("auto_fix", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "ERROR_RECOVERY"); val != "" {
		cm.SetConfig("error_recovery", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "ERROR_LOGGING"); val != "" {
		cm.SetConfig("error_logging", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "ERROR_METRICS"); val != "" {
		cm.SetConfig("error_metrics", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "ENABLE_CACHING"); val != "" {
		cm.SetConfig("enable_caching", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "ENABLE_COMPRESSION"); val != "" {
		cm.SetConfig("enable_compression", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "ENABLE_ENCRYPTION"); val != "" {
		cm.SetConfig("enable_encryption", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "WPS_COMPATIBILITY"); val != "" {
		cm.SetConfig("wps_compatibility", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "OFFICE_COMPATIBILITY"); val != "" {
		cm.SetConfig("office_compatibility", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "DEBUG_MODE"); val != "" {
		cm.SetConfig("debug_mode", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "VERBOSE_LOGGING"); val != "" {
		cm.SetConfig("verbose_logging", strings.ToLower(val) == "true")
	}
	if val := os.Getenv(ec.prefix + "PERFORMANCE_PROFILING"); val != "" {
		cm.SetConfig("performance_profiling", strings.ToLower(val) == "true")
	}
}

// ConfigValidator validates configuration values
type ConfigValidator struct {
	config *LibraryConfig
}

// NewConfigValidator creates a new configuration validator
func NewConfigValidator(config *LibraryConfig) *ConfigValidator {
	return &ConfigValidator{
		config: config,
	}
}

// Validate validates the configuration
func (cv *ConfigValidator) Validate() []string {
	var errors []string
	
	// Validate general settings
	if cv.config.Language == "" {
		errors = append(errors, "language cannot be empty")
	}
	if cv.config.DefaultEncoding == "" {
		errors = append(errors, "default_encoding cannot be empty")
	}
	
	// Validate performance settings
	if cv.config.MaxFileSize <= 0 {
		errors = append(errors, "max_file_size must be positive")
	}
	if cv.config.MemoryLimit <= 0 {
		errors = append(errors, "memory_limit must be positive")
	}
	if cv.config.Timeout <= 0 {
		errors = append(errors, "timeout must be positive")
	}
	if cv.config.BufferSize <= 0 {
		errors = append(errors, "buffer_size must be positive")
	}
	
	// Validate document settings
	if cv.config.DefaultFont == "" {
		errors = append(errors, "default_font cannot be empty")
	}
	if cv.config.DefaultFontSize <= 0 {
		errors = append(errors, "default_font_size must be positive")
	}
	
	// Validate format settings
	if len(cv.config.SupportedFormats) == 0 {
		errors = append(errors, "supported_formats cannot be empty")
	}
	
	// Validate validation settings
	if len(cv.config.ValidationRules) == 0 && cv.config.EnableValidation {
		errors = append(errors, "validation_rules cannot be empty when validation is enabled")
	}
	
	return errors
} 