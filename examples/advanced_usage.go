package main

import (
	"fmt"
	"log"
	"os"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func main() {
	fmt.Println("=== Go OpenXML SDK 高级使用示例 ===")
	
	// 示例1：基本文档操作
	basicDocumentExample()
	
	// 示例2：段落和格式分析
	paragraphAnalysisExample()
	
	// 示例3：表格处理
	tableAnalysisExample()
	
	fmt.Println("=== 示例完成 ===")
}

func basicDocumentExample() {
	fmt.Println("\n--- 基本文档操作 ---")
	
	filename := "example.docx"
	
	// 检查文件是否存在
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		fmt.Printf("文件 %s 不存在，跳过基本文档示例\n", filename)
		return
	}
	
	// 打开文档
	doc, err := wordprocessingml.Open(filename)
	if err != nil {
		log.Printf("无法打开文档 %s: %v", filename, err)
		return
	}
	defer doc.Close()
	
	// 获取文档文本
	text, err := doc.GetText()
	if err != nil {
		log.Printf("无法获取文档文本: %v", err)
		return
	}
	
	fmt.Printf("文档文本长度: %d 字符\n", len(text))
	if len(text) > 100 {
		fmt.Printf("文档预览: %s...\n", text[:100])
	} else {
		fmt.Printf("文档内容: %s\n", text)
	}
}

func paragraphAnalysisExample() {
	fmt.Println("\n--- 段落和格式分析 ---")
	
	filename := "example.docx"
	
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		fmt.Printf("文件 %s 不存在，跳过段落分析示例\n", filename)
		return
	}
	
	doc, err := wordprocessingml.Open(filename)
	if err != nil {
		log.Printf("无法打开文档 %s: %v", filename, err)
		return
	}
	defer doc.Close()
	
	// 获取段落
	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		log.Printf("无法获取段落: %v", err)
		return
	}
	
	fmt.Printf("文档包含 %d 个段落\n", len(paragraphs))
	
	// 分析段落格式
	for i, paragraph := range paragraphs {
		fmt.Printf("段落 %d:\n", i+1)
		fmt.Printf("  文本: %s\n", paragraph.Text)
		fmt.Printf("  样式: %s\n", paragraph.Style)
		fmt.Printf("  运行数: %d\n", len(paragraph.Runs))
		
		// 分析运行格式
		for j, run := range paragraph.Runs {
			fmt.Printf("    运行 %d: '%s'\n", j+1, run.Text)
			if run.Bold {
				fmt.Printf("      粗体: 是\n")
			}
			if run.Italic {
				fmt.Printf("      斜体: 是\n")
			}
			if run.Underline {
				fmt.Printf("      下划线: 是\n")
			}
			if run.FontSize > 0 {
				fmt.Printf("      字体大小: %d\n", run.FontSize)
			}
			if run.FontName != "" {
				fmt.Printf("      字体: %s\n", run.FontName)
			}
		}
		fmt.Println()
	}
}

func tableAnalysisExample() {
	fmt.Println("\n--- 表格处理 ---")
	
	filename := "example.docx"
	
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		fmt.Printf("文件 %s 不存在，跳过表格分析示例\n", filename)
		return
	}
	
	doc, err := wordprocessingml.Open(filename)
	if err != nil {
		log.Printf("无法打开文档 %s: %v", filename, err)
		return
	}
	defer doc.Close()
	
	// 获取表格
	tables, err := doc.GetTables()
	if err != nil {
		log.Printf("无法获取表格: %v", err)
		return
	}
	
	fmt.Printf("文档包含 %d 个表格\n", len(tables))
	
	// 分析表格结构
	for i, table := range tables {
		fmt.Printf("表格 %d:\n", i+1)
		fmt.Printf("  行数: %d\n", len(table.Rows))
		fmt.Printf("  列数: %d\n", table.Columns)
		
		// 显示表格内容
		for j, row := range table.Rows {
			fmt.Printf("  行 %d:\n", j+1)
			for k, cell := range row.Cells {
				fmt.Printf("    单元格 %d: %s\n", k+1, cell.Text)
			}
		}
		fmt.Println()
	}
}

// 辅助函数：创建示例文档
func createExampleDocument() {
	fmt.Println("\n--- 创建示例文档 ---")
	
	// 这里可以添加创建示例文档的代码
	// 在实际实现中，我们会提供文档创建功能
	
	fmt.Println("文档创建功能将在后续版本中实现")
} 