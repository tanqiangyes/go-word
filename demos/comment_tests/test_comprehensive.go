package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("测试文件生成功能...")

	// 测试1: 使用DocumentWriter
	fmt.Println("\n1. 测试DocumentWriter...")
	testDocumentWriter()

	fmt.Println("\n测试完成！请检查生成的文件是否可以正常打开。")
}

func testDocumentWriter() {
	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加段落
	err = docWriter.AddParagraph("Hello, World! 这是一个测试文档。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph:", err)
	}

	err = docWriter.AddParagraph("第二个段落，包含中文测试。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph:", err)
	}

	// 添加格式化段落
	formattedRuns := []types.Run{
		{
			Text:     "这是",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "粗体",
			Bold:     true,
			FontName: "宋体",
			FontSize: 14,
		},
		{
			Text:     "文本。",
			FontName: "宋体",
			FontSize: 12,
		},
	}
	
	err = docWriter.AddFormattedParagraph("这是粗体文本。", "Normal", formattedRuns)
	if err != nil {
		log.Fatal("Failed to add formatted paragraph:", err)
	}

	// 添加表格
	tableData := [][]string{
		{"姓名", "年龄", "职业"},
		{"张三", "25", "工程师"},
		{"李四", "30", "设计师"},
		{"王五", "28", "产品经理"},
	}
	
	err = docWriter.AddTable(tableData)
	if err != nil {
		log.Fatal("Failed to add table:", err)
	}

	// 添加更多段落
	err = docWriter.AddParagraph("表格测试完成。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph:", err)
	}

	// 保存文档
	filename := "test_comprehensive.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("DocumentWriter测试完成，文件已保存: %s\n", filename)
	fmt.Println("请尝试打开生成的文件，检查是否可以正常显示。")
}
