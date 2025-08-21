package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🔧 开始 WPS 语法修复后测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加标题
	fmt.Println("1. 添加文档标题...")
	err = docWriter.AddParagraph("WPS 语法修复后测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	// 添加段落
	fmt.Println("2. 添加段落...")
	
	paragraphText := "这是第一个测试段落。"
	err = docWriter.AddParagraph(paragraphText, "Normal")
	if err != nil {
		log.Fatal("Failed to add first paragraph:", err)
	}

	// 添加批注
	fmt.Println("3. 添加批注...")
	err = docWriter.AddComment("测试员", "这是一个测试批注。", paragraphText)
	if err != nil {
		log.Fatal("Failed to add comment:", err)
	}
	
	err = docWriter.AddParagraph("这是第二个测试段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}

	// 保存文档
	filename := "wps_syntax_fixed_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n🎉 WPS 语法修复后测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 使用语法修复后的 DocumentWriter")
	fmt.Println("2. 基本的段落添加")
	fmt.Println("3. 添加了一个批注")
	
	fmt.Println("\n🔧 语法修复内容：")
	fmt.Println("- 修复了所有样式标签的语法错误")
	fmt.Println("- 确保所有标签都正确关闭")
	fmt.Println("- 统一了标签格式")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 Word 中打开文档，检查批注是否显示")
	fmt.Println("- 在 WPS 中打开文档，检查批注是否显示")
	fmt.Println("- 检查是否还有\"样式1\"错误")
	
	fmt.Println("\n🏆 这是语法修复后的测试，应该能解决所有问题！")
}
