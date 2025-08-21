package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🔧 开始简化 WPS 兼容性批注测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加标题
	fmt.Println("1. 添加文档标题...")
	err = docWriter.AddParagraph("简化 WPS 批注测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	// 添加段落并添加批注
	fmt.Println("2. 添加段落并添加批注...")
	
	paragraphText := "这是一个测试段落。"
	err = docWriter.AddParagraph(paragraphText, "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph:", err)
	}
	
	// 添加批注
	err = docWriter.AddComment("测试员", "这是一个测试批注。", paragraphText)
	if err != nil {
		log.Fatal("Failed to add comment:", err)
	}
	fmt.Println("✅ 批注添加成功")

	// 添加第二个段落
	paragraphText2 := "这是第二个测试段落。"
	err = docWriter.AddParagraph(paragraphText2, "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}
	
	// 添加第二个批注
	err = docWriter.AddComment("审核员", "第二个批注测试。", paragraphText2)
	if err != nil {
		log.Fatal("Failed to add second comment:", err)
	}
	fmt.Println("✅ 第二个批注添加成功")

	// 保存文档
	filename := "wps_simple_comment_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n🎉 简化 WPS 兼容性批注测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 文档标题")
	fmt.Println("2. 第一个段落 + 测试员批注")
	fmt.Println("3. 第二个段落 + 审核员批注")
	
	fmt.Println("\n🔧 修复内容：")
	fmt.Println("- 简化了文档 XML 的命名空间")
	fmt.Println("- 移除了可能导致 WPS 无法打开的复杂命名空间")
	fmt.Println("- 保留了必要的批注功能")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 WPS 中打开文档")
	fmt.Println("- 点击 '审阅' 选项卡")
	fmt.Println("- 点击 '显示批注' 按钮")
	fmt.Println("- 批注应该显示在右侧边栏中")
	
	fmt.Println("\n🏆 如果这个简化版本能在 WPS 中打开，说明问题在于命名空间！")
}
