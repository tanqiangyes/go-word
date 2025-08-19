package examples

import (
	"context"
	"fmt"
	"strings"
	"time"

	"github.com/tanqiangyes/go-word/pkg/plugin"
)

// TextFormatterPlugin 文本格式化插件
type TextFormatterPlugin struct {
	config map[string]interface{}
}

// NewTextFormatterPlugin 创建文本格式化插件
func NewTextFormatterPlugin() *TextFormatterPlugin {
	return &TextFormatterPlugin{}
}

// GetInfo 获取插件信息
func (tfp *TextFormatterPlugin) GetInfo() *plugin.PluginInfo {
	return &plugin.PluginInfo{
		ID:          "text_formatter",
		Name:        "文本格式化插件",
		Version:     "1.0.0",
		Description: "提供文本格式化功能，包括大小写转换、去除多余空格等",
		Author:      "Go Word Team",
		License:     "MIT",
		Category:    "text",
		Tags:        []string{"text", "format", "utility"},
		Required:    []string{},
		Optional:    []string{},
		CreatedAt:   time.Now(),
		UpdatedAt:   time.Now(),
		Metadata: map[string]interface{}{
			"supported_formats": []string{"uppercase", "lowercase", "titlecase", "trim", "normalize"},
		},
	}
}

// Initialize 初始化插件
func (tfp *TextFormatterPlugin) Initialize(config map[string]interface{}) error {
	tfp.config = config
	return nil
}

// Execute 执行插件
func (tfp *TextFormatterPlugin) Execute(ctx context.Context, args map[string]interface{}) (*plugin.PluginResult, error) {
	text, ok := args["text"].(string)
	if !ok {
		return &plugin.PluginResult{
			Success: false,
			Message: "缺少文本参数",
		}, fmt.Errorf("缺少文本参数")
	}
	
	format, ok := args["format"].(string)
	if !ok {
		format = "normalize" // 默认格式
	}
	
	var result string
	var err error
	
	switch format {
	case "uppercase":
		result = strings.ToUpper(text)
	case "lowercase":
		result = strings.ToLower(text)
	case "titlecase":
		result = strings.Title(strings.ToLower(text))
	case "trim":
		result = strings.TrimSpace(text)
	case "normalize":
		result = tfp.normalizeText(text)
	default:
		return &plugin.PluginResult{
			Success: false,
			Message: fmt.Sprintf("不支持的格式: %s", format),
		}, fmt.Errorf("不支持的格式: %s", format)
	}
	
	return &plugin.PluginResult{
		Success: true,
		Data: map[string]interface{}{
			"original_text": text,
			"formatted_text": result,
			"format":        format,
		},
		Message: "文本格式化成功",
	}, nil
}

// Cleanup 清理插件资源
func (tfp *TextFormatterPlugin) Cleanup() error {
	tfp.config = nil
	return nil
}

// normalizeText 标准化文本
func (tfp *TextFormatterPlugin) normalizeText(text string) string {
	// 去除多余空格
	text = strings.TrimSpace(text)
	
	// 将多个空格替换为单个空格
	text = strings.Join(strings.Fields(text), " ")
	
	// 标准化标点符号
	text = strings.ReplaceAll(text, "  ", " ")
	text = strings.ReplaceAll(text, " ,", ",")
	text = strings.ReplaceAll(text, " .", ".")
	text = strings.ReplaceAll(text, " !", "!")
	text = strings.ReplaceAll(text, " ?", "?")
	
	return text
}
