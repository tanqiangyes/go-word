package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🔧 开始增强 WPS 兼容性批注测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加标题
	fmt.Println("1. 添加文档标题...")
	err = docWriter.AddParagraph("增强 WPS 兼容性批注测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	// 添加段落并添加批注
	fmt.Println("2. 添加段落并添加批注...")
	
	paragraphText := "这是一个测试段落，用于验证 WPS 批注功能。"
	err = docWriter.AddParagraph(paragraphText, "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph:", err)
	}
	
	// 添加批注
	err = docWriter.AddComment("测试员", "这是一个测试批注，用于验证 WPS 兼容性。", paragraphText)
	if err != nil {
		log.Fatal("Failed to add comment:", err)
	}
	fmt.Println("✅ 批注添加成功")

	// 添加第二个段落
	paragraphText2 := "这是第二个测试段落，包含更多内容。"
	err = docWriter.AddParagraph(paragraphText2, "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}
	
	// 添加第二个批注
	err = docWriter.AddComment("审核员", "第二个批注测试，用于验证 WPS 中的显示效果。", paragraphText2)
	if err != nil {
		log.Fatal("Failed to add second comment:", err)
	}
	fmt.Println("✅ 第二个批注添加成功")

	// 添加第三个段落
	paragraphText3 := "这是第三个测试段落，用于最终验证。"
	err = docWriter.AddParagraph(paragraphText3, "Normal")
	if err != nil {
		log.Fatal("Failed to add third paragraph:", err)
	}
	
	// 添加第三个批注
	err = docWriter.AddComment("管理员", "第三个批注测试，确保 WPS 兼容性。", paragraphText3)
	if err != nil {
		log.Fatal("Failed to add third comment:", err)
	}
	fmt.Println("✅ 第三个批注添加成功")

	// 保存文档
	filename := "wps_enhanced_comment_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n🎉 增强 WPS 兼容性批注测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 文档标题")
	fmt.Println("2. 第一个段落 + 测试员批注")
	fmt.Println("3. 第二个段落 + 审核员批注")
	fmt.Println("4. 第三个段落 + 管理员批注")
	
	fmt.Println("\n🔧 修复内容：")
	fmt.Println("- 修复了样式 XML 中的语法错误")
	fmt.Println("- 改进了批注 XML 结构")
	fmt.Println("- 添加了更多 WPS 兼容性属性")
	fmt.Println("- 使用正确的命名空间和标签")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 WPS 中打开文档")
	fmt.Println("- 点击 '审阅' 选项卡")
	fmt.Println("- 点击 '显示批注' 按钮")
	fmt.Println("- 批注应该显示在右侧边栏中")
	fmt.Println("- Word 中不应该再出现'样式1'错误")
	
	fmt.Println("\n🏆 如果批注能在 WPS 中正常显示且 Word 无错误，说明所有问题都已修复！")
}
