package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🔧 开始文件句柄修复测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加内容
	fmt.Println("1. 添加文档内容...")
	err = docWriter.AddParagraph("文件句柄修复测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	err = docWriter.AddParagraph("这是一个测试文档，用于验证文件句柄问题是否已修复。", "Normal")
	if err != nil {
		log.Fatal("Failed to add description:", err)
	}

	// 保存文档
	filename := "file_handle_fix_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("✅ 文档已保存: %s\n", filename)

	// 测试打开和修改文档
	fmt.Println("\n2. 测试打开和修改文档...")
	
	// 创建新的文档写入器来打开刚保存的文档
	modifyWriter := writer.NewDocumentWriter()
	
	// 打开文档进行修改
	err = modifyWriter.OpenForModification(filename)
	if err != nil {
		log.Fatal("Failed to open document for modification:", err)
	}
	
	fmt.Println("✅ 文档打开成功，没有文件句柄错误")
	
	// 添加新内容
	err = modifyWriter.AddParagraph("这是通过OpenForModification添加的新段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add new paragraph:", err)
	}
	
	// 保存修改后的文档
	modifiedFilename := "file_handle_fix_modified.docx"
	err = modifyWriter.Save(modifiedFilename)
	if err != nil {
		log.Fatal("Failed to save modified document:", err)
	}
	
	fmt.Printf("✅ 修改后的文档已保存: %s\n", modifiedFilename)
	
	fmt.Println("\n🎉 文件句柄修复测试完成！")
	fmt.Println("📋 测试结果：")
	fmt.Println("1. 创建新文档 ✅")
	fmt.Println("2. 保存文档 ✅")
	fmt.Println("3. 打开文档进行修改 ✅")
	fmt.Println("4. 修改并保存文档 ✅")
	
	fmt.Println("\n🔧 修复内容：")
	fmt.Println("- 修复了OPC容器中的文件句柄管理问题")
	fmt.Println("- 确保文件句柄在正确的时机关闭")
	fmt.Println("- 解决了'file already closed'错误")
	
	fmt.Println("\n🏆 现在应该可以正常打开和修改DOCX文件了！")
}
