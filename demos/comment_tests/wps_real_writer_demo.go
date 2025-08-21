package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🔧 开始 WPS 实际 DocumentWriter 兼容性测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加标题
	fmt.Println("1. 添加文档标题...")
	err = docWriter.AddParagraph("WPS 实际 DocumentWriter 测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	// 添加段落
	fmt.Println("2. 添加段落...")
	
	err = docWriter.AddParagraph("这是第一个测试段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add first paragraph:", err)
	}
	
	err = docWriter.AddParagraph("这是第二个测试段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}

	// 保存文档
	filename := "wps_real_writer_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n🎉 WPS 实际 DocumentWriter 兼容性测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 使用实际的 DocumentWriter")
	fmt.Println("2. 基本的段落添加")
	fmt.Println("3. 没有任何批注功能")
	
	fmt.Println("\n🔧 测试目的：")
	fmt.Println("- 验证我们实际的 DocumentWriter 是否能在 WPS 中打开")
	fmt.Println("- 如果这个文档能打开，说明问题在于批注功能")
	fmt.Println("- 如果这个文档不能打开，说明问题在于 DocumentWriter 本身")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 WPS 中打开文档")
	fmt.Println("- 检查是否能正常显示文本内容")
	fmt.Println("- 检查样式是否正常")
	
	fmt.Println("\n🏆 这是使用实际代码的测试，帮助我们找到真正的问题！")
}
