package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("=== WPS 批注兼容性测试 ===")

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

	// 添加批注 - 使用英文作者名（WPS 可能对中文支持更好）
	err = docWriter.AddComment("John Smith", "This paragraph contains important information that needs highlighting.", "这是第一个段落，包含重要信息。")
	if err != nil {
		log.Printf("添加批注失败: %v", err)
		return
	}

	err = docWriter.AddComment("Jane Doe", "Suggest adding more detailed information here.", "这是第二个段落，需要审查。")
	if err != nil {
		log.Printf("添加批注失败: %v", err)
		return
	}

	err = docWriter.AddComment("Mike Johnson", "This paragraph content is good and can be kept.", "这是第三个段落，最终内容。")
	if err != nil {
		log.Printf("添加批注失败: %v", err)
		return
	}

	// 保存文档
	err = docWriter.Save("wps_comments_test.docx")
	if err != nil {
		log.Printf("保存文档失败: %v", err)
		return
	}

	fmt.Println("✅ 文档已保存为 wps_comments_test.docx")
	fmt.Println()
	fmt.Println("📋 文档包含以下批注:")
	fmt.Println("   - John Smith: This paragraph contains important information that needs highlighting.")
	fmt.Println("   - Jane Doe: Suggest adding more detailed information here.")
	fmt.Println("   - Mike Johnson: This paragraph content is good and can be kept.")
	fmt.Println()
	fmt.Println("💡 WPS 测试说明:")
	fmt.Println("   1. 在 WPS Office 中打开 wps_comments_test.docx")
	fmt.Println("   2. 点击 '审阅' 选项卡")
	fmt.Println("   3. 点击 '显示批注' 或 '批注' 按钮")
	fmt.Println("   4. 如果 WPS 中能看到批注，说明兼容性问题已解决")
	fmt.Println("   5. 如果还是看不到，可能需要检查 WPS 的批注显示设置")
	fmt.Println()
	fmt.Println("🔧 技术改进:")
	fmt.Println("   - 添加了 w:initials 属性")
	fmt.Println("   - 改进了批注文本的格式设置")
	fmt.Println("   - 使用标准的 rId1 关系 ID")
	fmt.Println("   - 使用英文作者名进行测试")
} 