package wordprocessingml

import (
	"testing"
)

// TestNewFileEmbedder 测试创建文件嵌入器
func TestNewFileEmbedder(t *testing.T) {
	// 创建默认配置
	config := &FileEmbedderConfig{
		MaxFileSize:       1024 * 1024, // 1MB
		MaxTotalSize:      10 * 1024 * 1024, // 10MB
		AllowedMimeTypes:  []string{"text/plain", "image/jpeg", "image/png"},
		AllowedExtensions: []string{".txt", ".jpg", ".png"},
		CompressionLevel:  6,
		EncryptionEnabled: false,
		CacheEnabled:      true,
		CacheTTL:          3600,
		ValidateChecksums: true,
	}

	// 创建文档（模拟）
	doc := &Document{}

	// 测试创建文件嵌入器
	embedder := NewFileEmbedder(doc, config)
	if embedder == nil {
		t.Fatal("文件嵌入器创建失败")
	}

	// 验证配置
	if embedder.config.MaxFileSize != 1024*1024 {
		t.Errorf("期望最大文件大小为1MB，实际为%d", embedder.config.MaxFileSize)
	}

	if embedder.config.MaxTotalSize != 10*1024*1024 {
		t.Errorf("期望最大总大小为10MB，实际为%d", embedder.config.MaxTotalSize)
	}

	if len(embedder.config.AllowedMimeTypes) != 3 {
		t.Errorf("期望允许的MIME类型数量为3，实际为%d", len(embedder.config.AllowedMimeTypes))
	}

	if len(embedder.config.AllowedExtensions) != 3 {
		t.Errorf("期望允许的扩展名数量为3，实际为%d", len(embedder.config.AllowedExtensions))
	}
}

// TestNewFileEmbedderWithConfig 测试使用配置创建文件嵌入器
func TestNewFileEmbedderWithConfig(t *testing.T) {
	config := &FileEmbedderConfig{
		MaxFileSize:       512 * 1024, // 512KB
		MaxTotalSize:      5 * 1024 * 1024, // 5MB
		AllowedMimeTypes:  []string{"text/plain"},
		AllowedExtensions: []string{".txt"},
		CompressionLevel:  9,
		EncryptionEnabled: true,
		CacheEnabled:      false,
		CacheTTL:          1800,
		ValidateChecksums: false,
	}

	// 创建文档（模拟）
	doc := &Document{}

	embedder := NewFileEmbedder(doc, config)
	if embedder == nil {
		t.Fatal("文件嵌入器创建失败")
	}

	if embedder.config.MaxFileSize != 512*1024 {
		t.Errorf("配置最大文件大小不匹配，期望: 512KB, 实际: %d", embedder.config.MaxFileSize)
	}

	if embedder.config.MaxTotalSize != 5*1024*1024 {
		t.Errorf("配置最大总大小不匹配，期望: 5MB, 实际: %d", embedder.config.MaxTotalSize)
	}

	if len(embedder.config.AllowedMimeTypes) != 1 {
		t.Errorf("配置允许的MIME类型数量不匹配，期望: 1, 实际: %d", len(embedder.config.AllowedMimeTypes))
	}

	if len(embedder.config.AllowedExtensions) != 1 {
		t.Errorf("配置允许的扩展名数量不匹配，期望: 1, 实际: %d", len(embedder.config.AllowedExtensions))
	}
}

// TestNewFileEmbedderWithNilConfig 测试使用nil配置创建文件嵌入器
func TestNewFileEmbedderWithNilConfig(t *testing.T) {
	// 创建文档（模拟）
	doc := &Document{}

	// 测试使用nil配置创建
	embedder := NewFileEmbedder(doc, nil)
	if embedder == nil {
		t.Fatal("使用nil配置创建文件嵌入器失败")
	}

	// 验证使用默认配置
	if embedder.config == nil {
		t.Error("应该使用默认配置")
	}
}
