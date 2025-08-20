package word

import (
    "context"
    "fmt"
    "os"
    "sync"

    "github.com/tanqiangyes/go-word/pkg/utils"
)

// ImageProcessor 图片处理器
type ImageProcessor struct {
    images  map[string]*ImageProcessorImage
    effects map[string]*ImageProcessorEffect
    formats map[string]*ImageProcessorFormat
    mu      sync.RWMutex
    logger  *utils.Logger
    config  *ImageProcessorConfig
}

// ImageProcessorImage 图片
type ImageProcessorImage struct {
    ID        string                    `json:"id"`
    Path      string                    `json:"path"`
    Data      []byte                    `json:"data"`
    Format    ImageProcessorImageFormat `json:"format"`
    Size      *ImageProcessorSize       `json:"size"`
    Position  *ImageProcessorPosition   `json:"position"`
    Effects   []*ImageProcessorEffect   `json:"effects"`
    Metadata  map[string]interface{}    `json:"metadata"`
    CreatedAt int64                     `json:"created_at"`
    UpdatedAt int64                     `json:"updated_at"`
}

// ImageProcessorSize 图片尺寸
type ImageProcessorSize struct {
    Width               int     `json:"width"`
    Height              int     `json:"height"`
    ScaleX              float64 `json:"scale_x"`
    ScaleY              float64 `json:"scale_y"`
    MaintainAspectRatio bool    `json:"maintain_aspect_ratio"`
}

// ImageProcessorPosition 图片位置
type ImageProcessorPosition struct {
    X         float64                 `json:"x"`
    Y         float64                 `json:"y"`
    ZIndex    int                     `json:"z_index"`
    Alignment ImageProcessorAlignment `json:"alignment"`
    Wrapping  ImageProcessorWrapping  `json:"wrapping"`
}

// ImageProcessorEffect 图片效果
type ImageProcessorEffect struct {
    ID         string                   `json:"id"`
    Type       ImageProcessorEffectType `json:"type"`
    Parameters map[string]interface{}   `json:"parameters"`
    Intensity  float64                  `json:"intensity"`
    Enabled    bool                     `json:"enabled"`
}

// ImageProcessorFormat 图片格式
type ImageProcessorFormat struct {
    Name        string `json:"name"`
    Extension   string `json:"extension"`
    MimeType    string `json:"mime_type"`
    Quality     int    `json:"quality"`
    Compression int    `json:"compression"`
}

// ImageProcessorConfig 配置
type ImageProcessorConfig struct {
    MaxImages        int      `json:"max_images"`
    MaxImageSize     int64    `json:"max_image_size"`
    SupportedFormats []string `json:"supported_formats"`
    DefaultQuality   int      `json:"default_quality"`
    AutoCompression  bool     `json:"auto_compression"`
    CacheEnabled     bool     `json:"cache_enabled"`
    CacheSize        int      `json:"cache_size"`
    TempDirectory    string   `json:"temp_directory"`
}

// 常量定义
const (
    // 图片格式
    ImageProcessorImageFormatJPEG ImageProcessorImageFormat = "jpeg"
    ImageProcessorImageFormatPNG  ImageProcessorImageFormat = "png"
    ImageProcessorImageFormatGIF  ImageProcessorImageFormat = "gif"
    ImageProcessorImageFormatBMP  ImageProcessorImageFormat = "bmp"
    ImageProcessorImageFormatTIFF ImageProcessorImageFormat = "tiff"
    ImageProcessorImageFormatWEBP ImageProcessorImageFormat = "webp"

    // 对齐方式
    ImageProcessorAlignmentLeft   ImageProcessorAlignment = "left"
    ImageProcessorAlignmentCenter ImageProcessorAlignment = "center"
    ImageProcessorAlignmentRight  ImageProcessorAlignment = "right"
    ImageProcessorAlignmentTop    ImageProcessorAlignment = "top"
    ImageProcessorAlignmentMiddle ImageProcessorAlignment = "middle"
    ImageProcessorAlignmentBottom ImageProcessorAlignment = "bottom"

    // 环绕方式
    ImageProcessorWrappingInline  ImageProcessorWrapping = "inline"
    ImageProcessorWrappingSquare  ImageProcessorWrapping = "square"
    ImageProcessorWrappingTight   ImageProcessorWrapping = "tight"
    ImageProcessorWrappingThrough ImageProcessorWrapping = "through"
    ImageProcessorWrappingBehind  ImageProcessorWrapping = "behind"
    ImageProcessorWrappingInFront ImageProcessorWrapping = "in_front"

    // 效果类型
    ImageProcessorEffectTypeBrightness   ImageProcessorEffectType = "brightness"
    ImageProcessorEffectTypeContrast     ImageProcessorEffectType = "contrast"
    ImageProcessorEffectTypeSaturation   ImageProcessorEffectType = "saturation"
    ImageProcessorEffectTypeBlur         ImageProcessorEffectType = "blur"
    ImageProcessorEffectTypeSharpen      ImageProcessorEffectType = "sharpen"
    ImageProcessorEffectTypeGrayscale    ImageProcessorEffectType = "grayscale"
    ImageProcessorEffectTypeSepia        ImageProcessorEffectType = "sepia"
    ImageProcessorEffectTypeInvert       ImageProcessorEffectType = "invert"
    ImageProcessorEffectTypeTransparency ImageProcessorEffectType = "transparency"
    ImageProcessorEffectTypeBorder       ImageProcessorEffectType = "border"
    ImageProcessorEffectTypeShadow       ImageProcessorEffectType = "shadow"
    ImageProcessorEffectTypeReflection   ImageProcessorEffectType = "reflection"
)

// 类型定义
type ImageProcessorImageFormat string
type ImageProcessorAlignment string
type ImageProcessorWrapping string
type ImageProcessorEffectType string

// NewImageProcessor 创建新的图片处理器
func NewImageProcessor(config *ImageProcessorConfig) *ImageProcessor {
    if config == nil {
        config = &ImageProcessorConfig{
            MaxImages:        1000,
            MaxImageSize:     50 * 1024 * 1024, // 50MB
            SupportedFormats: []string{"jpeg", "png", "gif", "bmp", "tiff", "webp"},
            DefaultQuality:   85,
            AutoCompression:  true,
            CacheEnabled:     true,
            CacheSize:        100,
            TempDirectory:    "/tmp",
        }
    }

    ip := &ImageProcessor{
        images:  make(map[string]*ImageProcessorImage),
        effects: make(map[string]*ImageProcessorEffect),
        formats: make(map[string]*ImageProcessorFormat),
        config:  config,
        logger:  utils.NewLogger(utils.LogLevelInfo, nil),
    }

    // 初始化支持的格式
    ip.initializeFormats()

    return ip
}

// LoadImage 加载图片
func (ip *ImageProcessor) LoadImage(ctx context.Context, path string) (*ImageProcessorImage, error) {
    ip.mu.Lock()
    defer ip.mu.Unlock()

    // 检查文件是否存在
    if _, err := os.Stat(path); os.IsNotExist(err) {
        return nil, utils.NewStructuredDocumentError(utils.ErrFileNotFound, fmt.Sprintf("图片文件不存在: %s", path))
    }

    // 检查文件大小
    fileInfo, err := os.Stat(path)
    if err != nil {
        return nil, utils.NewStructuredDocumentError(utils.ErrFilePermission, fmt.Sprintf("无法访问文件: %s", path))
    }

    if fileInfo.Size() > ip.config.MaxImageSize {
        return nil, utils.NewStructuredDocumentError(utils.ErrFileTooLarge, fmt.Sprintf("图片文件过大: %d bytes", fileInfo.Size()))
    }

    // 读取文件数据
    data, err := os.ReadFile(path)
    if err != nil {
        return nil, utils.NewStructuredDocumentError(utils.ErrFileCorrupted, fmt.Sprintf("无法读取文件: %s", path))
    }

    // 检测图片格式
    format := ip.detectImageFormat(data)
    if format == "" {
        return nil, utils.NewStructuredDocumentError(utils.ErrFormatUnsupported, "不支持的图片格式")
    }

    // 获取图片尺寸
    size, err := ip.getImageSize(data, format)
    if err != nil {
        return nil, utils.NewStructuredDocumentError(utils.ErrContentInvalid, fmt.Sprintf("无法获取图片尺寸: %v", err))
    }

    // 创建图片对象
    image := &ImageProcessorImage{
        ID:     utils.GenerateID(),
        Path:   path,
        Data:   data,
        Format: format,
        Size:   size,
        Position: &ImageProcessorPosition{
            X:         0,
            Y:         0,
            ZIndex:    0,
            Alignment: ImageProcessorAlignmentLeft,
            Wrapping:  ImageProcessorWrappingInline,
        },
        Effects:   make([]*ImageProcessorEffect, 0),
        Metadata:  make(map[string]interface{}),
        CreatedAt: utils.GetCurrentTimestamp(),
        UpdatedAt: utils.GetCurrentTimestamp(),
    }

    // 存储图片
    ip.images[image.ID] = image

    ip.logger.Info("图片已加载，图片ID: %s, 路径: %s, 格式: %s, 尺寸: %v", image.ID, path, format, size)

    return image, nil
}

// InsertImage 插入图片
func (ip *ImageProcessor) InsertImage(ctx context.Context, imageID string, position *ImageProcessorPosition) error {
    ip.mu.Lock()
    defer ip.mu.Unlock()

    image, exists := ip.images[imageID]
    if !exists {
        return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图片不存在")
    }

    // 更新位置
    if position != nil {
        image.Position = position
    }

    image.UpdatedAt = utils.GetCurrentTimestamp()

    ip.logger.Info("图片已插入，图片ID: %s, 位置: %v", imageID, position)

    return nil
}

// ResizeImage 调整图片大小
func (ip *ImageProcessor) ResizeImage(ctx context.Context, imageID string, size *ImageProcessorSize) error {
    ip.mu.Lock()
    defer ip.mu.Unlock()

    image, exists := ip.images[imageID]
    if !exists {
        return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图片不存在")
    }

    // 更新尺寸
    if size != nil {
        image.Size = size
    }

    image.UpdatedAt = utils.GetCurrentTimestamp()

    ip.logger.Info("图片尺寸已调整，图片ID: %s, 尺寸: %v", imageID, size)

    return nil
}

// MoveImage 移动图片
func (ip *ImageProcessor) MoveImage(ctx context.Context, imageID string, x, y float64) error {
    ip.mu.Lock()
    defer ip.mu.Unlock()

    image, exists := ip.images[imageID]
    if !exists {
        return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图片不存在")
    }

    // 更新位置
    image.Position.X = x
    image.Position.Y = y
    image.UpdatedAt = utils.GetCurrentTimestamp()

    ip.logger.Info("图片已移动，图片ID: %s, X: %f, Y: %f", imageID, x, y)

    return nil
}

// ApplyEffect 应用效果
func (ip *ImageProcessor) ApplyEffect(ctx context.Context, imageID string, effect *ImageProcessorEffect) error {
    ip.mu.Lock()
    defer ip.mu.Unlock()

    image, exists := ip.images[imageID]
    if !exists {
        return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图片不存在")
    }

    // 生成效果ID
    if effect.ID == "" {
        effect.ID = utils.GenerateID()
    }

    // 应用效果
    image.Effects = append(image.Effects, effect)
    image.UpdatedAt = utils.GetCurrentTimestamp()

    // 存储效果
    ip.effects[effect.ID] = effect

    ip.logger.Info("效果已应用，图片ID: %s, 效果ID: %s, 效果类型: %s, 强度: %f", imageID, effect.ID, effect.Type, effect.Intensity)

    return nil
}

// RemoveEffect 移除效果
func (ip *ImageProcessor) RemoveEffect(ctx context.Context, imageID string, effectID string) error {
    ip.mu.Lock()
    defer ip.mu.Unlock()

    image, exists := ip.images[imageID]
    if !exists {
        return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图片不存在")
    }

    // 查找并移除效果
    for i, effect := range image.Effects {
        if effect.ID == effectID {
            image.Effects = append(image.Effects[:i], image.Effects[i+1:]...)
            delete(ip.effects, effectID)
            image.UpdatedAt = utils.GetCurrentTimestamp()

            ip.logger.Info("效果已移除，图片ID: %s, 效果ID: %s", imageID, effectID)

            return nil
        }
    }

    return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "效果不存在")
}

// ConvertFormat 转换格式
func (ip *ImageProcessor) ConvertFormat(ctx context.Context, imageID string, targetFormat ImageProcessorImageFormat, quality int) error {
    ip.mu.Lock()
    defer ip.mu.Unlock()

    image, exists := ip.images[imageID]
    if !exists {
        return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图片不存在")
    }

    // 检查目标格式是否支持
    if !ip.isFormatSupported(string(targetFormat)) {
        return utils.NewStructuredDocumentError(utils.ErrFormatUnsupported, fmt.Sprintf("不支持的格式: %s", targetFormat))
    }

    // 转换格式
    convertedData, err := ip.convertImageFormat(image.Data, image.Format, targetFormat, quality)
    if err != nil {
        return utils.NewStructuredDocumentError(utils.ErrFormatConversion, fmt.Sprintf("格式转换失败: %v", err))
    }

    // 更新图片数据
    image.Data = convertedData
    image.Format = targetFormat
    image.UpdatedAt = utils.GetCurrentTimestamp()

    ip.logger.Info("图片格式已转换，图片ID: %s, 旧格式: %s, 新格式: %s, 质量: %d", imageID, image.Format, targetFormat, quality)

    return nil
}

// GetImage 获取图片
func (ip *ImageProcessor) GetImage(imageID string) (*ImageProcessorImage, error) {
    ip.mu.RLock()
    defer ip.mu.RUnlock()

    image, exists := ip.images[imageID]
    if !exists {
        return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图片不存在")
    }

    return image, nil
}

// GetImageThumbnail 获取图片缩略图
func (ip *ImageProcessor) GetImageThumbnail(ctx context.Context, imageID string, width, height int) ([]byte, error) {
    image, err := ip.GetImage(imageID)
    if err != nil {
        return nil, err
    }

    // 生成缩略图
    thumbnailData, err := ip.generateThumbnail(image.Data, image.Format, width, height)
    if err != nil {
        return nil, utils.NewStructuredDocumentError(utils.ErrContentInvalid, fmt.Sprintf("生成缩略图失败: %v", err))
    }

    return thumbnailData, nil
}

// DeleteImage 删除图片
func (ip *ImageProcessor) DeleteImage(ctx context.Context, imageID string) error {
    ip.mu.Lock()
    defer ip.mu.Unlock()

    image, exists := ip.images[imageID]
    if !exists {
        return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "图片不存在")
    }

    // 删除图片
    delete(ip.images, imageID)

    // 删除相关效果
    for _, effect := range image.Effects {
        delete(ip.effects, effect.ID)
    }

    ip.logger.Info("图片已删除，图片ID: %s, 路径: %s", imageID, image.Path)

    return nil
}

// GetStats 获取统计信息
func (ip *ImageProcessor) GetStats() map[string]interface{} {
    ip.mu.RLock()
    defer ip.mu.RUnlock()

    stats := map[string]interface{}{
        "total_images":      len(ip.images),
        "total_effects":     len(ip.effects),
        "supported_formats": ip.config.SupportedFormats,
    }

    // 按格式统计
    formatCount := make(map[ImageProcessorImageFormat]int)
    for _, image := range ip.images {
        formatCount[image.Format]++
    }
    stats["format_count"] = formatCount

    // 按效果类型统计
    effectTypeCount := make(map[ImageProcessorEffectType]int)
    for _, effect := range ip.effects {
        effectTypeCount[effect.Type]++
    }
    stats["effect_type_count"] = effectTypeCount

    return stats
}

// 辅助方法

// initializeFormats 初始化支持的格式
func (ip *ImageProcessor) initializeFormats() {
    formats := map[string]*ImageProcessorFormat{
        "jpeg": {Name: "JPEG", Extension: ".jpg", MimeType: "image/jpeg", Quality: 85, Compression: 0},
        "png":  {Name: "PNG", Extension: ".png", MimeType: "image/png", Quality: 100, Compression: 0},
        "gif":  {Name: "GIF", Extension: ".gif", MimeType: "image/gif", Quality: 100, Compression: 0},
        "bmp":  {Name: "BMP", Extension: ".bmp", MimeType: "image/bmp", Quality: 100, Compression: 0},
        "tiff": {Name: "TIFF", Extension: ".tiff", MimeType: "image/tiff", Quality: 100, Compression: 0},
        "webp": {Name: "WebP", Extension: ".webp", MimeType: "image/webp", Quality: 85, Compression: 0},
    }

    for key, format := range formats {
        ip.formats[key] = format
    }
}

// detectImageFormat 检测图片格式
func (ip *ImageProcessor) detectImageFormat(data []byte) ImageProcessorImageFormat {
    if len(data) < 4 {
        return ""
    }

    // 检查文件头
    header := data[:4]

    switch {
    case len(header) >= 2 && header[0] == 0xFF && header[1] == 0xD8:
        return ImageProcessorImageFormatJPEG
    case len(header) >= 8 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47:
        return ImageProcessorImageFormatPNG
    case len(header) >= 6 && header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46:
        return ImageProcessorImageFormatGIF
    case len(header) >= 2 && header[0] == 0x42 && header[1] == 0x4D:
        return ImageProcessorImageFormatBMP
    case len(header) >= 4 && header[0] == 0x49 && header[1] == 0x49 && header[2] == 0x2A && header[3] == 0x00:
        return ImageProcessorImageFormatTIFF
    case len(header) >= 4 && header[0] == 0x52 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x46:
        // 检查WebP
        if len(data) >= 12 && data[8] == 0x57 && data[9] == 0x45 && data[10] == 0x42 && data[11] == 0x50 {
            return ImageProcessorImageFormatWEBP
        }
    }

    return ""
}

// getImageSize 获取图片尺寸
func (ip *ImageProcessor) getImageSize(data []byte, format ImageProcessorImageFormat) (*ImageProcessorSize, error) {
    // 这里应该实现真正的图片尺寸检测
    // 为了简化，我们返回默认尺寸
    return &ImageProcessorSize{
        Width:               800,
        Height:              600,
        ScaleX:              1.0,
        ScaleY:              1.0,
        MaintainAspectRatio: true,
    }, nil
}

// isFormatSupported 检查格式是否支持
func (ip *ImageProcessor) isFormatSupported(format string) bool {
    for _, supported := range ip.config.SupportedFormats {
        if supported == format {
            return true
        }
    }
    return false
}

// convertImageFormat 转换图片格式
func (ip *ImageProcessor) convertImageFormat(data []byte, fromFormat, toFormat ImageProcessorImageFormat, quality int) ([]byte, error) {
    // 这里应该实现真正的格式转换
    // 为了简化，我们返回原数据
    return data, nil
}

// generateThumbnail 生成缩略图
func (ip *ImageProcessor) generateThumbnail(data []byte, format ImageProcessorImageFormat, width, height int) ([]byte, error) {
    // 这里应该实现真正的缩略图生成
    // 为了简化，我们返回原数据
    return data, nil
}
