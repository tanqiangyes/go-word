package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("创建最小化测试文档...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 只添加一个简单段落
	err = docWriter.AddParagraph("这是一个最小化测试文档。", "")
	if err != nil {
		log.Fatal("Failed to add paragraph:", err)
	}

	// 保存文档
	filename := "minimal_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("最小化测试完成，文件已保存: %s\n", filename)
	fmt.Println("请尝试打开这个文件，看看是否可以正常显示。")
}
