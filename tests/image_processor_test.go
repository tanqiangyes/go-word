package tests

import (
    "context"
    "os"
    "testing"

    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// TestImageProcessor 测试图片处理器
func TestImageProcessor(t *testing.T) {
    // 创建图片处理器
    config := &word.ImageProcessorConfig{
        MaxImages:        100,
        MaxImageSize:     10 * 1024 * 1024, // 10MB
        SupportedFormats: []string{"jpeg", "png", "gif", "bmp", "tiff", "webp"},
        DefaultQuality:   85,
        AutoCompression:  true,
        CacheEnabled:     true,
        CacheSize:        50,
        TempDirectory:    "/tmp",
    }

    ip := word.NewImageProcessor(config)

    // 创建测试图片文件
    testImagePath := "/tmp/test_image.jpg"
    testImageData := []byte{0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01}

    err := os.WriteFile(testImagePath, testImageData, 0644)
    if err != nil {
        t.Fatalf("创建测试图片文件失败: %v", err)
    }
    defer os.Remove(testImagePath)

    // 测试加载图片
    image, err := ip.LoadImage(context.Background(), testImagePath)
    if err != nil {
        t.Fatalf("加载图片失败: %v", err)
    }

    if image.ID == "" {
        t.Error("图片ID为空")
    }

    if image.Path != testImagePath {
        t.Errorf("图片路径不匹配，期望: %s，实际: %s", testImagePath, image.Path)
    }

    if image.Format != word.ImageProcessorImageFormatJPEG {
        t.Errorf("图片格式不匹配，期望: jpeg，实际: %s", image.Format)
    }

    // 测试插入图片
    position := &word.ImageProcessorPosition{
        X:         100,
        Y:         200,
        ZIndex:    1,
        Alignment: word.ImageProcessorAlignmentCenter,
        Wrapping:  word.ImageProcessorWrappingSquare,
    }

    err = ip.InsertImage(context.Background(), image.ID, position)
    if err != nil {
        t.Fatalf("插入图片失败: %v", err)
    }

    // 测试调整图片大小
    size := &word.ImageProcessorSize{
        Width:               400,
        Height:              300,
        ScaleX:              0.8,
        ScaleY:              0.8,
        MaintainAspectRatio: true,
    }

    err = ip.ResizeImage(context.Background(), image.ID, size)
    if err != nil {
        t.Fatalf("调整图片大小失败: %v", err)
    }

    // 测试移动图片
    err = ip.MoveImage(context.Background(), image.ID, 150, 250)
    if err != nil {
        t.Fatalf("移动图片失败: %v", err)
    }

    // 测试应用效果
    effect := &word.ImageProcessorEffect{
        Type:       word.ImageProcessorEffectTypeBrightness,
        Parameters: map[string]interface{}{"value": 1.2},
        Intensity:  0.8,
        Enabled:    true,
    }

    err = ip.ApplyEffect(context.Background(), image.ID, effect)
    if err != nil {
        t.Fatalf("应用效果失败: %v", err)
    }

    // 测试获取图片
    retrievedImage, err := ip.GetImage(image.ID)
    if err != nil {
        t.Fatalf("获取图片失败: %v", err)
    }

    if retrievedImage.ID != image.ID {
        t.Errorf("图片ID不匹配，期望: %s，实际: %s", image.ID, retrievedImage.ID)
    }

    // 测试删除图片
    err = ip.DeleteImage(context.Background(), image.ID)
    if err != nil {
        t.Fatalf("删除图片失败: %v", err)
    }

    // 验证图片已删除
    _, err = ip.GetImage(image.ID)
    if err == nil {
        t.Error("图片应该已被删除")
    }

    // 测试获取统计信息
    stats := ip.GetStats()
    if stats["total_images"] != 0 {
        t.Errorf("图片数量不匹配，期望: 0，实际: %v", stats["total_images"])
    }
}

// TestImageProcessorErrorHandling 测试图片处理器错误处理
func TestImageProcessorErrorHandling(t *testing.T) {
    ip := word.NewImageProcessor(nil)

    // 测试加载不存在的文件
    _, err := ip.LoadImage(context.Background(), "/nonexistent/image.jpg")
    if err == nil {
        t.Error("应该返回错误，但未返回")
    }

    // 测试获取不存在的图片
    _, err = ip.GetImage("nonexistent")
    if err == nil {
        t.Error("应该返回错误，但未返回")
    }

    // 测试删除不存在的图片
    err = ip.DeleteImage(context.Background(), "nonexistent")
    if err == nil {
        t.Error("应该返回错误，但未返回")
    }

    // 测试插入不存在的图片
    err = ip.InsertImage(context.Background(), "nonexistent", nil)
    if err == nil {
        t.Error("应该返回错误，但未返回")
    }

    // 测试调整不存在的图片大小
    err = ip.ResizeImage(context.Background(), "nonexistent", nil)
    if err == nil {
        t.Error("应该返回错误，但未返回")
    }

    // 测试移动不存在的图片
    err = ip.MoveImage(context.Background(), "nonexistent", 0, 0)
    if err == nil {
        t.Error("应该返回错误，但未返回")
    }

    // 测试对不存在的图片应用效果
    effect := &word.ImageProcessorEffect{
        Type:       word.ImageProcessorEffectTypeBrightness,
        Parameters: map[string]interface{}{"value": 1.2},
        Intensity:  0.8,
        Enabled:    true,
    }

    err = ip.ApplyEffect(context.Background(), "nonexistent", effect)
    if err == nil {
        t.Error("应该返回错误，但未返回")
    }
}
