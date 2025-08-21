package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("开始文档修改功能测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加初始内容
	fmt.Println("1. 创建初始文档内容...")
	err = docWriter.AddParagraph("原始文档内容", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	err = docWriter.AddParagraph("这是一个测试文档，用于测试文本替换功能。", "Normal")
	if err != nil {
		log.Fatal("Failed to add description:", err)
	}

	err = docWriter.AddParagraph("客户姓名：{{客户姓名}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add customer name placeholder:", err)
	}

	err = docWriter.AddParagraph("订单编号：{{订单编号}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add order number placeholder:", err)
	}

	err = docWriter.AddParagraph("产品名称：{{产品名称}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add product name placeholder:", err)
	}

	err = docWriter.AddParagraph("订单金额：{{订单金额}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add order amount placeholder:", err)
	}

	// 保存原始文档
	originalFilename := "original_document.docx"
	err = docWriter.Save(originalFilename)
	if err != nil {
		log.Fatal("Failed to save original document:", err)
	}

	fmt.Printf("✅ 原始文档已保存: %s\n", originalFilename)

	// 测试文本替换功能
	fmt.Println("\n2. 测试文本替换功能...")
	
	// 替换客户信息
	err = docWriter.ReplaceText("{{客户姓名}}", "张三")
	if err != nil {
		log.Fatal("Failed to replace customer name:", err)
	}
	fmt.Println("✅ 替换客户姓名：{{客户姓名}} → 张三")

	err = docWriter.ReplaceText("{{订单编号}}", "ORD-2024-001")
	if err != nil {
		log.Fatal("Failed to replace order number:", err)
	}
	fmt.Println("✅ 替换订单编号：{{订单编号}} → ORD-2024-001")

	err = docWriter.ReplaceText("{{产品名称}}", "高性能笔记本电脑")
	if err != nil {
		log.Fatal("Failed to replace product name:", err)
	}
	fmt.Println("✅ 替换产品名称：{{产品名称}} → 高性能笔记本电脑")

	err = docWriter.ReplaceText("{{订单金额}}", "¥8,999")
	if err != nil {
		log.Fatal("Failed to replace order amount:", err)
	}
	fmt.Println("✅ 替换订单金额：{{订单金额}} → ¥8,999")

	// 保存修改后的文档
	modifiedFilename := "modified_document.docx"
	err = docWriter.Save(modifiedFilename)
	if err != nil {
		log.Fatal("Failed to save modified document:", err)
	}

	fmt.Printf("\n✅ 修改后的文档已保存: %s\n", modifiedFilename)

	// 测试多次替换
	fmt.Println("\n3. 测试多次文本替换...")
	
	// 添加更多内容
	err = docWriter.AddParagraph("", "Normal") // 空行
	if err != nil {
		log.Fatal("Failed to add empty line:", err)
	}

	err = docWriter.AddParagraph("订单状态：{{订单状态}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add order status:", err)
	}

	err = docWriter.AddParagraph("预计发货时间：{{发货时间}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add shipping time:", err)
	}

	err = docWriter.AddParagraph("配送地址：{{配送地址}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add delivery address:", err)
	}

	// 再次替换
	err = docWriter.ReplaceText("{{订单状态}}", "已确认")
	if err != nil {
		log.Fatal("Failed to replace order status:", err)
	}
	fmt.Println("✅ 替换订单状态：{{订单状态}} → 已确认")

	err = docWriter.ReplaceText("{{发货时间}}", "2024年1月25日")
	if err != nil {
		log.Fatal("Failed to replace shipping time:", err)
	}
	fmt.Println("✅ 替换发货时间：{{发货时间}} → 2024年1月25日")

	err = docWriter.ReplaceText("{{配送地址}}", "北京市朝阳区某某街道123号")
	if err != nil {
		log.Fatal("Failed to replace delivery address:", err)
	}
	fmt.Println("✅ 替换配送地址：{{配送地址}} → 北京市朝阳区某某街道123号")

	// 保存最终文档
	finalFilename := "final_document.docx"
	err = docWriter.Save(finalFilename)
	if err != nil {
		log.Fatal("Failed to save final document:", err)
	}

	fmt.Printf("\n✅ 最终文档已保存: %s\n", finalFilename)

	fmt.Println("\n📋 测试总结：")
	fmt.Printf("1. %s - 包含占位符的原始文档\n", originalFilename)
	fmt.Printf("2. %s - 部分替换后的文档\n", modifiedFilename)
	fmt.Printf("3. %s - 完全替换后的最终文档\n", finalFilename)
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 所有占位符是否被正确替换")
	fmt.Println("- 文本替换是否保持原有格式")
	fmt.Println("- 多次替换是否正常工作")
	fmt.Println("- 中文字符是否正确显示")
	fmt.Println("- 文档结构是否完整")
	
	fmt.Println("\n🎯 如果所有功能都正常，说明文本替换功能已经完善！")
}
