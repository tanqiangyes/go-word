package main

import (
	"context"
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/plugin"
	"github.com/tanqiangyes/go-word/pkg/plugin/examples"
)

// main 主函数 - 插件系统示例
func main() {
	fmt.Println("启动 Go Word 插件系统示例...")
	
	// 创建插件管理器
	pm := plugin.NewPluginManager()
	if pm == nil {
		log.Fatal("插件管理器创建失败")
	}
	
	// 创建文本格式化插件
	textFormatter := examples.NewTextFormatterPlugin()
	if textFormatter == nil {
		log.Fatal("文本格式化插件创建失败")
	}
	
	// 注册插件
	fmt.Println("正在注册文本格式化插件...")
	if err := pm.RegisterPlugin(textFormatter); err != nil {
		log.Fatalf("插件注册失败: %v", err)
	}
	fmt.Println("文本格式化插件注册成功")
	
	// 配置插件
	fmt.Println("正在配置插件...")
	config := map[string]interface{}{
		"default_format": "normalize",
		"enable_logging": true,
	}
	
	if err := pm.ConfigurePlugin("text_formatter", config); err != nil {
		log.Fatalf("插件配置失败: %v", err)
	}
	fmt.Println("插件配置成功")
	
	// 列出所有插件
	fmt.Println("\n已注册的插件:")
	plugins := pm.ListPlugins()
	for _, p := range plugins {
		fmt.Printf("  - %s (%s): %s\n", p.Name, p.Version, p.Description)
	}
	
	// 执行插件 - 大写转换
	fmt.Println("\n执行插件 - 大写转换:")
	ctx := context.Background()
	args := map[string]interface{}{
		"text":   "hello world",
		"format": "uppercase",
	}
	
	result, err := pm.ExecutePlugin(ctx, "text_formatter", args)
	if err != nil {
		log.Printf("插件执行失败: %v", err)
	} else {
		fmt.Printf("  原始文本: %s\n", args["text"])
		fmt.Printf("  格式化后: %s\n", result.Data["formatted_text"])
		fmt.Printf("  执行时间: %v\n", result.ExecutionTime)
	}
	
	// 执行插件 - 文本标准化
	fmt.Println("\n执行插件 - 文本标准化:")
	args = map[string]interface{}{
		"text":   "  hello   world  ,  how   are   you  ?  ",
		"format": "normalize",
	}
	
	result, err = pm.ExecutePlugin(ctx, "text_formatter", args)
	if err != nil {
		log.Printf("插件执行失败: %v", err)
	} else {
		fmt.Printf("  原始文本: %s\n", args["text"])
		fmt.Printf("  标准化后: %s\n", result.Data["formatted_text"])
		fmt.Printf("  执行时间: %v\n", result.ExecutionTime)
	}
	
	// 获取插件信息
	fmt.Println("\n插件详细信息:")
	info, err := pm.GetPluginInfo("text_formatter")
	if err != nil {
		log.Printf("获取插件信息失败: %v", err)
	} else {
		fmt.Printf("  ID: %s\n", info.ID)
		fmt.Printf("  名称: %s\n", info.Name)
		fmt.Printf("  版本: %s\n", info.Version)
		fmt.Printf("  描述: %s\n", info.Description)
		fmt.Printf("  作者: %s\n", info.Author)
		fmt.Printf("  类别: %s\n", info.Category)
		fmt.Printf("  标签: %v\n", info.Tags)
	}
	
	// 获取执行结果
	fmt.Println("\n插件执行结果:")
	execResult, err := pm.GetPluginResult("text_formatter")
	if err != nil {
		log.Printf("获取执行结果失败: %v", err)
	} else {
		fmt.Printf("  成功: %t\n", execResult.Success)
		fmt.Printf("  消息: %s\n", execResult.Message)
		fmt.Printf("  时间戳: %v\n", execResult.Timestamp)
	}
	
	// 获取插件指标
	fmt.Println("\n插件系统指标:")
	metrics := pm.GetMetrics()
	fmt.Printf("  总插件数: %d\n", metrics.TotalPlugins)
	fmt.Printf("  活跃插件数: %d\n", metrics.ActivePlugins)
	fmt.Printf("  总执行次数: %d\n", metrics.TotalExecutions)
	fmt.Printf("  成功次数: %d\n", metrics.SuccessCount)
	fmt.Printf("  错误次数: %d\n", metrics.ErrorCount)
	fmt.Printf("  平均执行时间: %v\n", metrics.AverageTime)
	fmt.Printf("  最后执行时间: %v\n", metrics.LastExecution)
	
	// 注销插件
	fmt.Println("\n正在注销插件...")
	if err := pm.UnregisterPlugin("text_formatter"); err != nil {
		log.Printf("插件注销失败: %v", err)
	} else {
		fmt.Println("插件注销成功")
	}
	
	fmt.Println("\n插件系统示例完成")
}
