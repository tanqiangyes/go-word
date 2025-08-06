package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("=== WPS 增强兼容性测试 ===")

	docWriter := writer.NewDocumentWriter()

	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Printf("创建文档失败: %v", err)
		return
	}

	// 添加段落
	err = docWriter.AddParagraph("这是第一个段落，包含重要信息。", "Normal")
	if err != nil {
		log.Printf("添加段落失败: %v", err)
		return
	}

	err = docWriter.AddParagraph("这是第二个段落，需要审查。", "Normal")
	if err != nil {
		log.Printf("添加段落失败: %v", err)
		return
	}

	err = docWriter.AddParagraph("这是第三个段落，最终内容。", "Normal")
	if err != nil {
		log.Printf("添加段落失败: %v", err)
		return
	}

	// 添加批注 - 使用不同的作者名格式
	err = docWriter.AddComment("张三", "这个段落包含重要信息，需要高亮显示。", "这是第一个段落，包含重要信息。")
	if err != nil {
		log.Printf("添加批注失败: %v", err)
		return
	}

	err = docWriter.AddComment("李四", "建议在这里添加更多详细信息。", "这是第二个段落，需要审查。")
	if err != nil {
		log.Printf("添加批注失败: %v", err)
		return
	}

	err = docWriter.AddComment("王五", "这个段落内容很好，可以保留。", "这是第三个段落，最终内容。")
	if err != nil {
		log.Printf("添加批注失败: %v", err)
		return
	}

	// 保存文档
	err = docWriter.Save("wps_enhanced_test.docx")
	if err != nil {
		log.Printf("保存文档失败: %v", err)
		return
	}

	fmt.Println("✅ 文档已保存为 wps_enhanced_test.docx")
	fmt.Println()
	fmt.Println("📋 文档包含以下批注:")
	fmt.Println("   - 张三: 这个段落包含重要信息，需要高亮显示")
	fmt.Println("   - 李四: 建议在这里添加更多详细信息")
	fmt.Println("   - 王五: 这个段落内容很好，可以保留")
	fmt.Println()
	fmt.Println("💡 WPS 测试说明:")
	fmt.Println("   1. 在 WPS Office 中打开 wps_enhanced_test.docx")
	fmt.Println("   2. 点击 '审阅' 选项卡")
	fmt.Println("   3. 点击 '显示批注' 或 '批注' 按钮")
	fmt.Println("   4. 如果 WPS 中能看到批注，说明兼容性问题已解决")
	fmt.Println("   5. 如果还是看不到，请尝试以下步骤:")
	fmt.Println("      - 检查 WPS 的批注显示设置")
	fmt.Println("      - 确保 WPS 版本支持批注功能")
	fmt.Println("      - 尝试在 WPS 中手动启用批注显示")
	fmt.Println()
	fmt.Println("🔧 技术改进:")
	fmt.Println("   - 添加了 w:initials 属性")
	fmt.Println("   - 改进了批注文本的格式设置")
	fmt.Println("   - 使用标准的 rId1 关系 ID")
	fmt.Println("   - 包含完整的段落和运行属性")
	fmt.Println("   - 支持中文作者名")
	fmt.Println()
	fmt.Println("📝 故障排除:")
	fmt.Println("   如果 WPS 中仍然看不到批注:")
	fmt.Println("   1. 确保 WPS 版本支持批注功能")
	fmt.Println("   2. 检查 WPS 的审阅设置")
	fmt.Println("   3. 尝试在 WPS 中手动显示批注")
	fmt.Println("   4. 确认文档没有损坏")
} 