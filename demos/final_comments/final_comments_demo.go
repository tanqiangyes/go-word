package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("=== 最终批注功能演示 ===")

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

	// 添加批注 - 使用中文作者名
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
	err = docWriter.Save("final_comments_demo.docx")
	if err != nil {
		log.Printf("保存文档失败: %v", err)
		return
	}

	fmt.Println("✅ 文档已保存为 final_comments_demo.docx")
	fmt.Println()
	fmt.Println("📋 文档包含以下批注:")
	fmt.Println("   - 张三: 这个段落包含重要信息，需要高亮显示")
	fmt.Println("   - 李四: 建议在这里添加更多详细信息")
	fmt.Println("   - 王五: 这个段落内容很好，可以保留")
	fmt.Println()
	fmt.Println("💡 测试说明:")
	fmt.Println("   1. 在 Microsoft Word 中打开 final_comments_demo.docx")
	fmt.Println("   2. 点击 '审阅' 选项卡")
	fmt.Println("   3. 点击 '显示批注' 按钮")
	fmt.Println("   4. 应该能在右侧边栏中看到批注")
	fmt.Println("   5. 如果 Word 中能看到批注，说明格式正确")
	fmt.Println("   6. 然后在 WPS 中测试同样的文档")
	fmt.Println()
	fmt.Println("🔧 技术特点:")
	fmt.Println("   - 基于 Open-XML-SDK 结构设计")
	fmt.Println("   - 使用 CommentManager 统一管理批注")
	fmt.Println("   - 支持中文作者名处理")
	fmt.Println("   - 生成完整的 OpenXML 结构")
	fmt.Println("   - 包含正确的批注范围标记和引用")
	fmt.Println("   - 生成正确的 relationships 和 content types")
} 