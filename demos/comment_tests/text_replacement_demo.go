package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("开始文本替换功能测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 添加包含占位符的段落
	fmt.Println("1. 添加包含占位符的段落...")
	err = docWriter.AddParagraph("尊敬的 {{客户姓名}}，您好！", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph with placeholder:", err)
	}

	err = docWriter.AddParagraph("感谢您选择我们的服务。您的订单号是：{{订单号}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}

	err = docWriter.AddParagraph("订单详情：", "Normal")
	if err != nil {
		log.Fatal("Failed to add order details header:", err)
	}

	// 添加表格
	fmt.Println("2. 添加订单表格...")
	orderTable := [][]string{
		{"商品名称", "数量", "单价", "小计"},
		{"{{商品1名称}}", "{{商品1数量}}", "{{商品1单价}}", "{{商品1小计}}"},
		{"{{商品2名称}}", "{{商品2数量}}", "{{商品2单价}}", "{{商品2小计}}"},
		{"{{商品3名称}}", "{{商品3数量}}", "{{商品3单价}}", "{{商品3小计}}"},
	}
	
	err = docWriter.AddTable(orderTable)
	if err != nil {
		log.Fatal("Failed to add order table:", err)
	}

	// 添加更多段落
	err = docWriter.AddParagraph("订单总金额：{{订单总金额}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add total amount paragraph:", err)
	}

	err = docWriter.AddParagraph("预计发货时间：{{发货时间}}", "Normal")
	if err != nil {
		log.Fatal("Failed to add shipping time paragraph:", err)
	}

	err = docWriter.AddParagraph("如有任何问题，请联系我们的客服团队。", "Normal")
	if err != nil {
		log.Fatal("Failed to add contact paragraph:", err)
	}

	// 保存原始文档
	originalFilename := "template_with_placeholders.docx"
	err = docWriter.Save(originalFilename)
	if err != nil {
		log.Fatal("Failed to save template document:", err)
	}

	fmt.Printf("✅ 模板文档已保存: %s\n", originalFilename)

	// 现在测试文本替换功能
	fmt.Println("\n3. 测试文本替换功能...")
	
	// 替换客户信息
	err = docWriter.ReplaceText("{{客户姓名}}", "张三")
	if err != nil {
		log.Fatal("Failed to replace customer name:", err)
	}

	err = docWriter.ReplaceText("{{订单号}}", "ORD-2024-001")
	if err != nil {
		log.Fatal("Failed to replace order number:", err)
	}

	// 替换商品信息
	err = docWriter.ReplaceText("{{商品1名称}}", "笔记本电脑")
	if err != nil {
		log.Fatal("Failed to replace product 1 name:", err)
	}

	err = docWriter.ReplaceText("{{商品1数量}}", "1")
	if err != nil {
		log.Fatal("Failed to replace product 1 quantity:", err)
	}

	err = docWriter.ReplaceText("{{商品1单价}}", "¥5999")
	if err != nil {
		log.Fatal("Failed to replace product 1 price:", err)
	}

	err = docWriter.ReplaceText("{{商品1小计}}", "¥5999")
	if err != nil {
		log.Fatal("Failed to replace product 1 subtotal:", err)
	}

	err = docWriter.ReplaceText("{{商品2名称}}", "无线鼠标")
	if err != nil {
		log.Fatal("Failed to replace product 2 name:", err)
	}

	err = docWriter.ReplaceText("{{商品2数量}}", "2")
	if err != nil {
		log.Fatal("Failed to replace product 2 quantity:", err)
	}

	err = docWriter.ReplaceText("{{商品2单价}}", "¥99")
	if err != nil {
		log.Fatal("Failed to replace product 2 price:", err)
	}

	err = docWriter.ReplaceText("{{商品2小计}}", "¥198")
	if err != nil {
		log.Fatal("Failed to replace product 2 subtotal:", err)
	}

	err = docWriter.ReplaceText("{{商品3名称}}", "机械键盘")
	if err != nil {
		log.Fatal("Failed to replace product 3 name:", err)
	}

	err = docWriter.ReplaceText("{{商品3数量}}", "1")
	if err != nil {
		log.Fatal("Failed to replace product 3 quantity:", err)
	}

	err = docWriter.ReplaceText("{{商品3单价}}", "¥299")
	if err != nil {
		log.Fatal("Failed to replace product 3 price:", err)
	}

	err = docWriter.ReplaceText("{{商品3小计}}", "¥299")
	if err != nil {
		log.Fatal("Failed to replace product 3 subtotal:", err)
	}

	// 替换其他信息
	err = docWriter.ReplaceText("{{订单总金额}}", "¥6496")
	if err != nil {
		log.Fatal("Failed to replace total amount:", err)
	}

	err = docWriter.ReplaceText("{{发货时间}}", "2024年1月20日")
	if err != nil {
		log.Fatal("Failed to replace shipping time:", err)
	}

	// 保存替换后的文档
	finalFilename := "filled_order_document.docx"
	err = docWriter.Save(finalFilename)
	if err != nil {
		log.Fatal("Failed to save filled document:", err)
	}

	fmt.Printf("✅ 填充后的文档已保存: %s\n", finalFilename)
	fmt.Println("\n请检查以下文件：")
	fmt.Printf("1. %s - 包含占位符的模板文档\n", originalFilename)
	fmt.Printf("2. %s - 填充了实际数据的文档\n", finalFilename)
	fmt.Println("\n验证要点：")
	fmt.Println("- 所有占位符是否被正确替换")
	fmt.Println("- 表格数据是否正确显示")
	fmt.Println("- 文档格式是否保持一致")
	fmt.Println("- 中文字符是否正确显示")
}
