package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🔧 开始最基础 WPS 兼容性测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加标题
	fmt.Println("1. 添加文档标题...")
	err = docWriter.AddParagraph("最基础 WPS 测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	// 添加段落
	fmt.Println("2. 添加段落...")
	
	err = docWriter.AddParagraph("这是一个测试段落，没有任何批注。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph:", err)
	}
	
	err = docWriter.AddParagraph("这是第二个测试段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}

	err = docWriter.AddParagraph("这是第三个测试段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add third paragraph:", err)
	}

	// 保存文档
	filename := "wps_basic_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n🎉 最基础 WPS 兼容性测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 文档标题")
	fmt.Println("2. 三个普通段落")
	fmt.Println("3. 没有任何批注功能")
	
	fmt.Println("\n🔧 测试目的：")
	fmt.Println("- 验证最基本的文档结构是否能在 WPS 中打开")
	fmt.Println("- 如果这个文档能打开，说明问题在于批注功能")
	fmt.Println("- 如果这个文档不能打开，说明问题在于基础文档结构")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 WPS 中打开文档")
	fmt.Println("- 检查是否能正常显示文本内容")
	
	fmt.Println("\n🏆 这是最基础的测试，如果连这个都打不开，说明问题很严重！")
}
