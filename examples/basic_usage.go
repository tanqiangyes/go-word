package main

import (
	"fmt"
	"log"
	
	"github.com/go-word/pkg/wordprocessingml"
)

func main() {
	// 示例：打开Word文档
	fmt.Println("=== Go OpenXML SDK 基本使用示例 ===")
	
	// 注意：这里需要一个实际的.docx文件来测试
	// 在实际使用中，请替换为您的文档路径
	filename := "example.docx"
	
	// 打开文档
	doc, err := wordprocessingml.Open(filename)
	if err != nil {
		log.Printf("无法打开文档 %s: %v", filename, err)
		fmt.Println("请确保有一个有效的.docx文件用于测试")
		return
	}
	defer doc.Close()
	
	// 获取文档文本
	text, err := doc.GetText()
	if err != nil {
		log.Printf("无法获取文档文本: %v", err)
		return
	}
	
	fmt.Printf("文档内容:\n%s\n", text)
	
	// 获取段落
	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		log.Printf("无法获取段落: %v", err)
		return
	}
	
	fmt.Printf("文档包含 %d 个段落\n", len(paragraphs))
	
	// 获取表格
	tables, err := doc.GetTables()
	if err != nil {
		log.Printf("无法获取表格: %v", err)
		return
	}
	
	fmt.Printf("文档包含 %d 个表格\n", len(tables))
	
	fmt.Println("=== 示例完成 ===")
} 