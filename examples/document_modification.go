package examples

import (
	"fmt"
	"log"
	"os"

	"github.com/tanqiangyes/go-word/pkg/writer"
	"github.com/tanqiangyes/go-word/pkg/types"
)

func DemoDocumentModification() {
	fmt.Println("=== Go OpenXML SDK 文档修改示例 ===")

	// 示例1：创建新文档
	createNewDocumentExample()

	// 示例2：修改现有文档
	modifyExistingDocumentExample()

	// 示例3：高级格式修改
	advancedFormattingExample()

	fmt.Println("=== 示例完成 ===")
}

func createNewDocumentExample() {
	fmt.Println("\n--- 创建新文档 ---")

	writer := writer.NewDocumentWriter()

	// 创建新文档
	err := writer.CreateNewDocument()
	if err != nil {
		log.Printf("创建新文档失败: %v", err)
		return
	}

	// 添加段落
	err = writer.AddParagraph("这是一个新创建的文档", "Normal")
	if err != nil {
		log.Printf("添加段落失败: %v", err)
		return
	}

	err = writer.AddParagraph("这是第二个段落，包含一些基本文本。", "Normal")
	if err != nil {
		log.Printf("添加段落失败: %v", err)
		return
	}

	// 添加表格
	tableData := [][]string{
		{"姓名", "年龄", "职业"},
		{"张三", "25", "工程师"},
		{"李四", "30", "设计师"},
		{"王五", "28", "产品经理"},
	}

	err = writer.AddTable(tableData)
	if err != nil {
		log.Printf("添加表格失败: %v", err)
		return
	}

	// 保存文档
	err = writer.Save("new_document.docx")
	if err != nil {
		log.Printf("保存文档失败: %v", err)
		return
	}

	fmt.Println("新文档已创建: new_document.docx")
}

func modifyExistingDocumentExample() {
	fmt.Println("\n--- 修改现有文档 ---")

	// 检查示例文档是否存在
	filename := "example.docx"
	if !fileExists(filename) {
		fmt.Printf("文件 %s 不存在，跳过修改示例\n", filename)
		return
	}

	// 打开现有文档进行修改
	writer := writer.NewDocumentWriter()
	err := writer.OpenForModification(filename)
	if err != nil {
		log.Printf("打开文档失败: %v", err)
		return
	}

	// 添加新段落
	err = writer.AddParagraph("这是通过程序添加的新段落。", "Normal")
	if err != nil {
		log.Printf("添加段落失败: %v", err)
		return
	}

	// 替换文本
	err = writer.ReplaceText("测试", "修改后的测试")
	if err != nil {
		log.Printf("替换文本失败: %v", err)
		return
	}

	// 保存修改后的文档
	err = writer.Save("modified_document.docx")
	if err != nil {
		log.Printf("保存文档失败: %v", err)
		return
	}

	fmt.Println("文档已修改并保存: modified_document.docx")
}

func advancedFormattingExample() {
	fmt.Println("\n--- 高级格式修改 ---")

	writer := writer.NewDocumentWriter()

	// 创建新文档
	err := writer.CreateNewDocument()
	if err != nil {
		log.Printf("创建新文档失败: %v", err)
		return
	}

	// 添加带格式的段落
	formattedRuns := []types.Run{
		{
			Text:     "这是",
			Bold:     true,
			FontSize: 16,
			FontName: "Arial",
		},
		{
			Text:     "一个",
			Italic:   true,
			FontSize: 14,
			FontName: "Times New Roman",
		},
		{
			Text:      "格式化的",
			Underline: true,
			FontSize:  12,
			FontName:  "Calibri",
		},
		{
			Text:     "段落。",
			FontSize: 12,
			FontName: "Arial",
		},
	}

	err = writer.AddFormattedParagraph("这是一个格式化的段落。", "Normal", formattedRuns)
	if err != nil {
		log.Printf("添加格式化段落失败: %v", err)
		return
	}

	// 添加简单段落
	err = writer.AddParagraph("这是另一个段落。", "Normal")
	if err != nil {
		log.Printf("添加段落失败: %v", err)
		return
	}

	// 保存文档
	err = writer.Save("formatted_document.docx")
	if err != nil {
		log.Printf("保存文档失败: %v", err)
		return
	}

	fmt.Println("格式化文档已创建: formatted_document.docx")
}

// 辅助函数：检查文件是否存在
func fileExists(filename string) bool {
	_, err := os.Stat(filename)
	return err == nil
} 