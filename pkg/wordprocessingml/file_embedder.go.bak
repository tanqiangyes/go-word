package wordprocessingml

import (
	"context"
	"crypto/md5"
	"encoding/hex"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/tanqiangyes/go-word/pkg/utils"
)

// FileEmbedder 文件嵌入器
type FileEmbedder struct {
	document      *Document
	embeddedFiles map[string]*EmbeddedFile
	linkRegistry  map[string]*DocumentLink
	config        *FileEmbedderConfig
	logger        *utils.Logger
	mu            sync.RWMutex
	metrics       *FileEmbedderMetrics
}

// EmbeddedFile 嵌入文件
type EmbeddedFile struct {
	ID           string                 `json:"id"`
	Name         string                 `json:"name"`
	OriginalPath string                 `json:"original_path"`
	MimeType     string                 `json:"mime_type"`
	Size         int64                  `json:"size"`
	Checksum     string                 `json:"checksum"`
	Data         []byte                 `json:"data,omitempty"`
	EmbedType    EmbedType              `json:"embed_type"`
	Position     *EmbedPosition         `json:"position"`
	Properties   map[string]interface{} `json:"properties"`
	CreatedAt    time.Time              `json:"created_at"`
	UpdatedAt    time.Time              `json:"updated_at"`
}

// DocumentLink 文档链接
type DocumentLink struct {
	ID          string                 `json:"id"`
	Type        LinkType               `json:"type"`
	Target      string                 `json:"target"`
	DisplayText string                 `json:"display_text"`
	Tooltip     string                 `json:"tooltip"`
	Style       *LinkStyle             `json:"style"`
	Position    *LinkPosition          `json:"position"`
	Properties  map[string]interface{} `json:"properties"`
	CreatedAt   time.Time              `json:"created_at"`
	UpdatedAt   time.Time              `json:"updated_at"`
}

// FileEmbedderConfig 文件嵌入器配置
type FileEmbedderConfig struct {
	MaxFileSize       int64    `json:"max_file_size"`
	MaxTotalSize      int64    `json:"max_total_size"`
	AllowedMimeTypes  []string `json:"allowed_mime_types"`
	AllowedExtensions []string `json:"allowed_extensions"`
	CompressionLevel  int      `json:"compression_level"`
	EncryptionEnabled bool     `json:"encryption_enabled"`
	CacheEnabled      bool     `json:"cache_enabled"`
	CacheTTL          int64    `json:"cache_ttl"`
	ValidateChecksums bool     `json:"validate_checksums"`
}

// FileEmbedderMetrics 文件嵌入器指标
type FileEmbedderMetrics struct {
	TotalFiles       int64         `json:"total_files"`
	TotalSize        int64         `json:"total_size"`
	TotalLinks       int64         `json:"total_links"`
	EmbedOperations  int64         `json:"embed_operations"`
	LinkOperations   int64         `json:"link_operations"`
	AverageFileSize  int64         `json:"average_file_size"`
	LastOperation    time.Time     `json:"last_operation"`
	OperationTime    time.Duration `json:"operation_time"`
}

// EmbedType 嵌入类型
type EmbedType string

const (
	EmbedTypeInline     EmbedType = "inline"     // 内联嵌入
	EmbedTypeAttachment EmbedType = "attachment" // 附件
	EmbedTypeReference  EmbedType = "reference"  // 引用
	EmbedTypeLink       EmbedType = "link"       // 链接
)

// LinkType 链接类型
type LinkType string

const (
	LinkTypeInternal LinkType = "internal" // 内部链接
	LinkTypeExternal LinkType = "external" // 外部链接
	LinkTypeEmail    LinkType = "email"    // 邮件链接
	LinkTypeFile     LinkType = "file"     // 文件链接
	LinkTypeBookmark LinkType = "bookmark" // 书签链接
)

// EmbedPosition 嵌入位置
type EmbedPosition struct {
	ParagraphIndex int                    `json:"paragraph_index"`
	RunIndex       int                    `json:"run_index"`
	CharIndex      int                    `json:"char_index"`
	Anchor         string                 `json:"anchor"`
	Properties     map[string]interface{} `json:"properties"`
}

// LinkPosition 链接位置
type LinkPosition struct {
	Start      *EmbedPosition         `json:"start"`
	End        *EmbedPosition         `json:"end"`
	Properties map[string]interface{} `json:"properties"`
}

// LinkStyle 链接样式
type LinkStyle struct {
	Color           string `json:"color"`
	Underline       bool   `json:"underline"`
	HoverColor      string `json:"hover_color"`
	VisitedColor    string `json:"visited_color"`
	FontWeight      string `json:"font_weight"`
	TextDecoration  string `json:"text_decoration"`
}

// EmbedResult 嵌入结果
type EmbedResult struct {
	Success      bool          `json:"success"`
	FileID       string        `json:"file_id"`
	OriginalSize int64         `json:"original_size"`
	EmbeddedSize int64         `json:"embedded_size"`
	CompressionRatio float64   `json:"compression_ratio"`
	ProcessTime  time.Duration `json:"process_time"`
	Error        error         `json:"error,omitempty"`
}

// LinkResult 链接结果
type LinkResult struct {
	Success     bool          `json:"success"`
	LinkID      string        `json:"link_id"`
	Target      string        `json:"target"`
	ProcessTime time.Duration `json:"process_time"`
	Error       error         `json:"error,omitempty"`
}

// NewFileEmbedder 创建文件嵌入器
func NewFileEmbedder(document *Document, config *FileEmbedderConfig) *FileEmbedder {
	if config == nil {
		config = getDefaultFileEmbedderConfig()
	}
	
	return &FileEmbedder{
		document:      document,
		embeddedFiles: make(map[string]*EmbeddedFile),
		linkRegistry:  make(map[string]*DocumentLink),
		config:        config,
		logger:        utils.NewLogger(utils.LogLevelInfo, os.Stdout),
		metrics:       &FileEmbedderMetrics{},
	}
}

// EmbedFile 嵌入文件
func (fe *FileEmbedder) EmbedFile(ctx context.Context, filePath string, embedType EmbedType, position *EmbedPosition) (*EmbedResult, error) {
	fe.mu.Lock()
	defer fe.mu.Unlock()
	
	startTime := time.Now()
	result := &EmbedResult{}
	
	// 验证文件
	fileInfo, err := os.Stat(filePath)
	if err != nil {
		result.Error = fmt.Errorf("file not found: %w", err)
		return result, result.Error
	}
	
	// 检查文件大小
	if fileInfo.Size() > fe.config.MaxFileSize {
		result.Error = fmt.Errorf("file size %d exceeds maximum allowed size %d", fileInfo.Size(), fe.config.MaxFileSize)
		return result, result.Error
	}
	
	// 检查总大小限制
	if fe.metrics.TotalSize+fileInfo.Size() > fe.config.MaxTotalSize {
		result.Error = fmt.Errorf("total embedded size would exceed maximum allowed size %d", fe.config.MaxTotalSize)
		return result, result.Error
	}
	
	// 验证文件类型
	if !fe.isFileTypeAllowed(filePath) {
		result.Error = fmt.Errorf("file type not allowed: %s", filepath.Ext(filePath))
		return result, result.Error
	}
	
	// 读取文件
	fileData, err := os.ReadFile(filePath)
	if err != nil {
		result.Error = fmt.Errorf("failed to read file: %w", err)
		return result, result.Error
	}
	
	// 计算校验和
	checksum := fe.calculateChecksum(fileData)
	
	// 检查是否已存在相同文件
	for _, existing := range fe.embeddedFiles {
		if existing.Checksum == checksum {
			fe.logger.Info("文件已存在，使用现有嵌入", map[string]interface{}{
				"file_path": filePath,
				"existing_id": existing.ID,
			})
			result.Success = true
			result.FileID = existing.ID
			result.OriginalSize = existing.Size
			result.EmbeddedSize = existing.Size
			result.ProcessTime = time.Since(startTime)
			return result, nil
		}
	}
	
	// 创建嵌入文件
	embeddedFile := &EmbeddedFile{
		ID:           utils.GenerateID(),
		Name:         filepath.Base(filePath),
		OriginalPath: filePath,
		MimeType:     fe.detectMimeType(filePath),
		Size:         fileInfo.Size(),
		Checksum:     checksum,
		Data:         fileData,
		EmbedType:    embedType,
		Position:     position,
		Properties:   make(map[string]interface{}),
		CreatedAt:    time.Now(),
		UpdatedAt:    time.Now(),
	}
	
	// 压缩数据（如果启用）
	if fe.config.CompressionLevel > 0 {
		compressedData, err := fe.compressData(fileData)
		if err != nil {
			fe.logger.Warning("文件压缩失败，使用原始数据", map[string]interface{}{
				"error": err.Error(),
			})
		} else {
			embeddedFile.Data = compressedData
			embeddedFile.Properties["compressed"] = true
			embeddedFile.Properties["original_size"] = fileInfo.Size()
		}
	}
	
	// 加密数据（如果启用）
	if fe.config.EncryptionEnabled {
		encryptedData, err := fe.encryptData(embeddedFile.Data)
		if err != nil {
			fe.logger.Warning("文件加密失败，使用原始数据", map[string]interface{}{
				"error": err.Error(),
			})
		} else {
			embeddedFile.Data = encryptedData
			embeddedFile.Properties["encrypted"] = true
		}
	}
	
	// 存储嵌入文件
	fe.embeddedFiles[embeddedFile.ID] = embeddedFile
	
	// 更新指标
	fe.updateEmbedMetrics(embeddedFile)
	
	// 设置结果
	result.Success = true
	result.FileID = embeddedFile.ID
	result.OriginalSize = fileInfo.Size()
	result.EmbeddedSize = int64(len(embeddedFile.Data))
	result.CompressionRatio = float64(result.EmbeddedSize) / float64(result.OriginalSize)
	result.ProcessTime = time.Since(startTime)
	
	fe.logger.Info("文件嵌入成功", map[string]interface{}{
		"file_id":          result.FileID,
		"file_path":        filePath,
		"original_size":    result.OriginalSize,
		"embedded_size":    result.EmbeddedSize,
		"compression_ratio": result.CompressionRatio,
		"process_time":     result.ProcessTime,
	})
	
	return result, nil
}

// CreateLink 创建链接
func (fe *FileEmbedder) CreateLink(ctx context.Context, linkType LinkType, target, displayText string, style *LinkStyle, position *LinkPosition) (*LinkResult, error) {
	fe.mu.Lock()
	defer fe.mu.Unlock()
	
	startTime := time.Now()
	result := &LinkResult{}
	
	// 验证链接目标
	if target == "" {
		result.Error = fmt.Errorf("link target cannot be empty")
		return result, result.Error
	}
	
	// 验证链接类型
	if err := fe.validateLinkTarget(linkType, target); err != nil {
		result.Error = err
		return result, result.Error
	}
	
	// 创建链接
	link := &DocumentLink{
		ID:          utils.GenerateID(),
		Type:        linkType,
		Target:      target,
		DisplayText: displayText,
		Style:       style,
		Position:    position,
		Properties:  make(map[string]interface{}),
		CreatedAt:   time.Now(),
		UpdatedAt:   time.Now(),
	}
	
	// 设置默认样式
	if link.Style == nil {
		link.Style = fe.getDefaultLinkStyle()
	}
	
	// 设置默认显示文本
	if link.DisplayText == "" {
		link.DisplayText = target
	}
	
	// 存储链接
	fe.linkRegistry[link.ID] = link
	
	// 更新指标
	fe.updateLinkMetrics()
	
	// 设置结果
	result.Success = true
	result.LinkID = link.ID
	result.Target = target
	result.ProcessTime = time.Since(startTime)
	
	fe.logger.Info("链接创建成功", map[string]interface{}{
		"link_id":      result.LinkID,
		"link_type":    linkType,
		"target":       target,
		"display_text": displayText,
		"process_time": result.ProcessTime,
	})
	
	return result, nil
}

// GetEmbeddedFile 获取嵌入文件
func (fe *FileEmbedder) GetEmbeddedFile(fileID string) (*EmbeddedFile, error) {
	fe.mu.RLock()
	defer fe.mu.RUnlock()
	
	file, exists := fe.embeddedFiles[fileID]
	if !exists {
		return nil, fmt.Errorf("embedded file not found: %s", fileID)
	}
	
	return file, nil
}

// GetLink 获取链接
func (fe *FileEmbedder) GetLink(linkID string) (*DocumentLink, error) {
	fe.mu.RLock()
	defer fe.mu.RUnlock()
	
	link, exists := fe.linkRegistry[linkID]
	if !exists {
		return nil, fmt.Errorf("link not found: %s", linkID)
	}
	
	return link, nil
}

// ExtractFile 提取嵌入文件
func (fe *FileEmbedder) ExtractFile(fileID, outputPath string) error {
	fe.mu.RLock()
	defer fe.mu.RUnlock()
	
	file, exists := fe.embeddedFiles[fileID]
	if !exists {
		return fmt.Errorf("embedded file not found: %s", fileID)
	}
	
	data := file.Data
	
	// 解密数据（如果需要）
	if encrypted, ok := file.Properties["encrypted"].(bool); ok && encrypted {
		decryptedData, err := fe.decryptData(data)
		if err != nil {
			return fmt.Errorf("failed to decrypt file: %w", err)
		}
		data = decryptedData
	}
	
	// 解压数据（如果需要）
	if compressed, ok := file.Properties["compressed"].(bool); ok && compressed {
		decompressedData, err := fe.decompressData(data)
		if err != nil {
			return fmt.Errorf("failed to decompress file: %w", err)
		}
		data = decompressedData
	}
	
	// 验证校验和（如果启用）
	if fe.config.ValidateChecksums {
		calculatedChecksum := fe.calculateChecksum(data)
		if calculatedChecksum != file.Checksum {
			return fmt.Errorf("checksum mismatch: expected %s, got %s", file.Checksum, calculatedChecksum)
		}
	}
	
	// 写入文件
	err := os.WriteFile(outputPath, data, 0644)
	if err != nil {
		return fmt.Errorf("failed to write file: %w", err)
	}
	
	fe.logger.Info("文件提取成功", map[string]interface{}{
		"file_id":     fileID,
		"output_path": outputPath,
		"size":        len(data),
	})
	
	return nil
}

// RemoveEmbeddedFile 移除嵌入文件
func (fe *FileEmbedder) RemoveEmbeddedFile(fileID string) error {
	fe.mu.Lock()
	defer fe.mu.Unlock()
	
	file, exists := fe.embeddedFiles[fileID]
	if !exists {
		return fmt.Errorf("embedded file not found: %s", fileID)
	}
	
	// 更新指标
	fe.metrics.TotalFiles--
	fe.metrics.TotalSize -= file.Size
	
	// 删除文件
	delete(fe.embeddedFiles, fileID)
	
	fe.logger.Info("嵌入文件已移除", map[string]interface{}{
		"file_id": fileID,
		"name":    file.Name,
	})
	
	return nil
}

// RemoveLink 移除链接
func (fe *FileEmbedder) RemoveLink(linkID string) error {
	fe.mu.Lock()
	defer fe.mu.Unlock()
	
	link, exists := fe.linkRegistry[linkID]
	if !exists {
		return fmt.Errorf("link not found: %s", linkID)
	}
	
	// 更新指标
	fe.metrics.TotalLinks--
	
	// 删除链接
	delete(fe.linkRegistry, linkID)
	
	fe.logger.Info("链接已移除", map[string]interface{}{
		"link_id": linkID,
		"target":  link.Target,
	})
	
	return nil
}

// GetMetrics 获取指标
func (fe *FileEmbedder) GetMetrics() *FileEmbedderMetrics {
	fe.mu.RLock()
	defer fe.mu.RUnlock()
	
	return fe.metrics
}

// ListEmbeddedFiles 列出所有嵌入文件
func (fe *FileEmbedder) ListEmbeddedFiles() []*EmbeddedFile {
	fe.mu.RLock()
	defer fe.mu.RUnlock()
	
	files := make([]*EmbeddedFile, 0, len(fe.embeddedFiles))
	for _, file := range fe.embeddedFiles {
		files = append(files, file)
	}
	
	return files
}

// ListLinks 列出所有链接
func (fe *FileEmbedder) ListLinks() []*DocumentLink {
	fe.mu.RLock()
	defer fe.mu.RUnlock()
	
	links := make([]*DocumentLink, 0, len(fe.linkRegistry))
	for _, link := range fe.linkRegistry {
		links = append(links, link)
	}
	
	return links
}

// 辅助方法

// isFileTypeAllowed 检查文件类型是否允许
func (fe *FileEmbedder) isFileTypeAllowed(filePath string) bool {
	ext := strings.ToLower(filepath.Ext(filePath))
	
	// 检查扩展名
	for _, allowed := range fe.config.AllowedExtensions {
		if ext == strings.ToLower(allowed) {
			return true
		}
	}
	
	// 检查MIME类型
	mimeType := fe.detectMimeType(filePath)
	for _, allowed := range fe.config.AllowedMimeTypes {
		if mimeType == allowed {
			return true
		}
	}
	
	return false
}

// detectMimeType 检测MIME类型
func (fe *FileEmbedder) detectMimeType(filePath string) string {
	ext := strings.ToLower(filepath.Ext(filePath))
	
	mimeTypes := map[string]string{
		".txt":  "text/plain",
		".pdf":  "application/pdf",
		".doc":  "application/msword",
		".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		".xls":  "application/vnd.ms-excel",
		".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		".ppt":  "application/vnd.ms-powerpoint",
		".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
		".jpg":  "image/jpeg",
		".jpeg": "image/jpeg",
		".png":  "image/png",
		".gif":  "image/gif",
		".svg":  "image/svg+xml",
		".zip":  "application/zip",
		".rar":  "application/x-rar-compressed",
		".7z":   "application/x-7z-compressed",
	}
	
	if mimeType, exists := mimeTypes[ext]; exists {
		return mimeType
	}
	
	return "application/octet-stream"
}

// calculateChecksum 计算校验和
func (fe *FileEmbedder) calculateChecksum(data []byte) string {
	hash := md5.Sum(data)
	return hex.EncodeToString(hash[:])
}

// compressData 压缩数据
func (fe *FileEmbedder) compressData(data []byte) ([]byte, error) {
	// 这里应该实现数据压缩逻辑
	// 为了简化，我们返回原始数据
	return data, nil
}

// decompressData 解压数据
func (fe *FileEmbedder) decompressData(data []byte) ([]byte, error) {
	// 这里应该实现数据解压逻辑
	// 为了简化，我们返回原始数据
	return data, nil
}

// encryptData 加密数据
func (fe *FileEmbedder) encryptData(data []byte) ([]byte, error) {
	// 这里应该实现数据加密逻辑
	// 为了简化，我们返回原始数据
	return data, nil
}

// decryptData 解密数据
func (fe *FileEmbedder) decryptData(data []byte) ([]byte, error) {
	// 这里应该实现数据解密逻辑
	// 为了简化，我们返回原始数据
	return data, nil
}

// validateLinkTarget 验证链接目标
func (fe *FileEmbedder) validateLinkTarget(linkType LinkType, target string) error {
	switch linkType {
	case LinkTypeExternal:
		if !strings.HasPrefix(target, "http://") && !strings.HasPrefix(target, "https://") {
			return fmt.Errorf("external link must start with http:// or https://")
		}
	case LinkTypeEmail:
		if !strings.Contains(target, "@") {
			return fmt.Errorf("email link must contain @ symbol")
		}
	case LinkTypeFile:
		if !filepath.IsAbs(target) && !strings.HasPrefix(target, "./") && !strings.HasPrefix(target, "../") {
			return fmt.Errorf("file link must be absolute or relative path")
		}
	}
	
	return nil
}

// getDefaultLinkStyle 获取默认链接样式
func (fe *FileEmbedder) getDefaultLinkStyle() *LinkStyle {
	return &LinkStyle{
		Color:          "#0066cc",
		Underline:      true,
		HoverColor:     "#004499",
		VisitedColor:   "#800080",
		FontWeight:     "normal",
		TextDecoration: "underline",
	}
}

// updateEmbedMetrics 更新嵌入指标
func (fe *FileEmbedder) updateEmbedMetrics(file *EmbeddedFile) {
	fe.metrics.TotalFiles++
	fe.metrics.TotalSize += file.Size
	fe.metrics.EmbedOperations++
	fe.metrics.LastOperation = time.Now()
	
	if fe.metrics.TotalFiles > 0 {
		fe.metrics.AverageFileSize = fe.metrics.TotalSize / fe.metrics.TotalFiles
	}
}

// updateLinkMetrics 更新链接指标
func (fe *FileEmbedder) updateLinkMetrics() {
	fe.metrics.TotalLinks++
	fe.metrics.LinkOperations++
	fe.metrics.LastOperation = time.Now()
}

// getDefaultFileEmbedderConfig 获取默认配置
func getDefaultFileEmbedderConfig() *FileEmbedderConfig {
	return &FileEmbedderConfig{
		MaxFileSize:      10 * 1024 * 1024, // 10MB
		MaxTotalSize:     100 * 1024 * 1024, // 100MB
		AllowedMimeTypes: []string{
			"text/plain",
			"application/pdf",
			"application/msword",
			"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			"image/jpeg",
			"image/png",
			"image/gif",
		},
		AllowedExtensions: []string{
			".txt", ".pdf", ".doc", ".docx",
			".jpg", ".jpeg", ".png", ".gif",
		},
		CompressionLevel:  0,
		EncryptionEnabled: false,
		CacheEnabled:      true,
		CacheTTL:          3600, // 1 hour
		ValidateChecksums: true,
	}
}