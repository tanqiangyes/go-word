package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("创建简单内容测试文档...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加简单的段落
	fmt.Println("1. 添加标题...")
	err = docWriter.AddParagraph("测试文档", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	fmt.Println("2. 添加段落1...")
	err = docWriter.AddParagraph("这是第一个段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph 1:", err)
	}

	fmt.Println("3. 添加段落2...")
	err = docWriter.AddParagraph("这是第二个段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph 2:", err)
	}

	fmt.Println("4. 添加段落3...")
	err = docWriter.AddParagraph("这是第三个段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph 3:", err)
	}

	fmt.Println("5. 添加段落4...")
	err = docWriter.AddParagraph("这是第四个段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph 4:", err)
	}

	fmt.Println("6. 添加段落5...")
	err = docWriter.AddParagraph("这是第五个段落。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph 5:", err)
	}

	// 保存文档
	filename := "simple_content_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("✅ 简单内容测试完成！文件已保存: %s\n", filename)
	fmt.Println("请打开这个文件，检查是否能看到以下内容：")
	fmt.Println("1. 测试文档")
	fmt.Println("2. 这是第一个段落。")
	fmt.Println("3. 这是第二个段落。")
	fmt.Println("4. 这是第三个段落。")
	fmt.Println("5. 这是第四个段落。")
	fmt.Println("6. 这是第五个段落。")
}
