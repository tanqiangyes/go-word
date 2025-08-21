package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🔧 开始 WPS 修复后批注功能兼容性测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加标题
	fmt.Println("1. 添加文档标题...")
	err = docWriter.AddParagraph("WPS 修复后批注功能测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	// 添加段落
	fmt.Println("2. 添加段落...")
	
	paragraphText := "这是第一个测试段落，包含重要信息。"
	err = docWriter.AddParagraph(paragraphText, "Normal")
	if err != nil {
		log.Fatal("Failed to add first paragraph:", err)
	}

	// 添加批注
	fmt.Println("3. 添加批注...")
	err = docWriter.AddComment("测试员", "这是一个测试批注，现在应该能在 WPS 中显示了！", paragraphText)
	if err != nil {
		log.Fatal("Failed to add comment:", err)
	}
	
	err = docWriter.AddParagraph("这是第二个测试段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}

	// 添加第二个批注
	paragraphText2 := "这是第三个测试段落，也包含重要信息。"
	err = docWriter.AddParagraph(paragraphText2, "Normal")
	if err != nil {
		log.Fatal("Failed to add third paragraph:", err)
	}

	err = docWriter.AddComment("审核员", "这是第二个测试批注，用于验证修复效果。", paragraphText2)
	if err != nil {
		log.Fatal("Failed to add second comment:", err)
	}

	// 保存文档
	filename := "wps_fixed_comments_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n🎉 WPS 修复后批注功能兼容性测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 使用修复后的 DocumentWriter")
	fmt.Println("2. 基本的段落添加")
	fmt.Println("3. 添加了两个批注")
	fmt.Println("4. 包含 WPS 兼容的样式定义")
	
	fmt.Println("\n🔧 修复内容：")
	fmt.Println("- 添加了 CommentReference 样式")
	fmt.Println("- 修复了批注引用的 XML 结构")
	fmt.Println("- 确保批注引用有正确的样式属性")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 WPS 中打开文档")
	fmt.Println("- 检查是否能正常显示文本内容")
	fmt.Println("- 检查批注是否能在 WPS 中显示")
	fmt.Println("- 检查批注引用是否有蓝色下划线")
	
	fmt.Println("\n🏆 这是修复后的测试，应该能解决 WPS 批注显示问题！")
}
