package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("开始全面功能测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 测试1: 基本段落
	fmt.Println("1. 测试基本段落...")
	err = docWriter.AddParagraph("这是一个功能测试文档", "Normal")
	if err != nil {
		log.Fatal("Failed to add basic paragraph:", err)
	}

	err = docWriter.AddParagraph("包含多种格式和内容的测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}

	// 测试2: 格式化段落
	fmt.Println("2. 测试格式化段落...")
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
			Text:     "文本，",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "斜体",
			Italic:   true,
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "文本，",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "下划线",
			Underline: true,
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "文本。",
			FontName: "宋体",
			FontSize: 12,
		},
	}
	
	err = docWriter.AddFormattedParagraph("这是格式化文本测试。", "Normal", formattedRuns)
	if err != nil {
		log.Fatal("Failed to add formatted paragraph:", err)
	}

	// 测试3: 不同字体的段落
	fmt.Println("3. 测试不同字体...")
	err = docWriter.AddParagraph("这是 Times New Roman 字体的测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add font test paragraph:", err)
	}

	// 测试4: 表格功能
	fmt.Println("4. 测试表格功能...")
	tableData := [][]string{
		{"产品名称", "价格", "库存", "状态"},
		{"笔记本电脑", "¥5999", "50", "有货"},
		{"无线鼠标", "¥99", "200", "有货"},
		{"机械键盘", "¥299", "30", "缺货"},
		{"显示器", "¥1299", "25", "有货"},
	}
	
	err = docWriter.AddTable(tableData)
	if err != nil {
		log.Fatal("Failed to add table:", err)
	}

	// 测试5: 更多段落内容
	fmt.Println("5. 测试更多段落内容...")
	err = docWriter.AddParagraph("表格测试完成。现在测试更多段落内容。", "Normal")
	if err != nil {
		log.Fatal("Failed to add paragraph after table:", err)
	}

	err = docWriter.AddParagraph("这是一个包含中文和英文混合的段落：Hello World! 你好世界！", "Normal")
	if err != nil {
		log.Fatal("Failed to add mixed language paragraph:", err)
	}

	err = docWriter.AddParagraph("测试特殊字符：@#$%^&*()_+-=[]{}|;':\",./<>?", "Normal")
	if err != nil {
		log.Fatal("Failed to add special characters paragraph:", err)
	}

	// 测试6: 长文本段落
	fmt.Println("6. 测试长文本段落...")
	longText := "这是一个很长的段落，用来测试文档处理长文本的能力。" +
		"长文本段落可以帮助我们验证文档的布局和格式是否正确。" +
		"当文本内容很长时，Word 会自动换行，我们需要确保换行后的格式保持一致。" +
		"这个测试段落包含了多个句子，涵盖了不同的标点符号和格式要求。"
	
	err = docWriter.AddParagraph(longText, "Normal")
	if err != nil {
		log.Fatal("Failed to add long text paragraph:", err)
	}

	// 测试7: 数字和列表样式
	fmt.Println("7. 测试数字和列表...")
	err = docWriter.AddParagraph("1. 第一项：基本功能测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add numbered item 1:", err)
	}

	err = docWriter.AddParagraph("2. 第二项：格式化功能测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add numbered item 2:", err)
	}

	err = docWriter.AddParagraph("3. 第三项：表格功能测试", "Normal")
	if err != nil {
		log.Fatal("Failed to add numbered item 3:", err)
	}

	// 测试8: 总结段落
	fmt.Println("8. 添加总结段落...")
	err = docWriter.AddParagraph("功能测试总结：", "Normal")
	if err != nil {
		log.Fatal("Failed to add summary header:", err)
	}

	summaryText := "本次测试涵盖了以下功能：基本段落创建、格式化文本（粗体、斜体、下划线）、" +
		"不同字体支持、表格创建、长文本处理、特殊字符支持、中英文混合等。" +
		"如果这个文档能够正常打开和显示，说明我们的 DOCX 生成功能已经基本完善。"
	
	err = docWriter.AddParagraph(summaryText, "Normal")
	if err != nil {
		log.Fatal("Failed to add summary text:", err)
	}

	// 保存文档
	filename := "feature_test_comprehensive.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n✅ 全面功能测试完成！文件已保存: %s\n", filename)
	fmt.Println("请打开这个文件，检查以下功能是否正确：")
	fmt.Println("1. 基本段落显示")
	fmt.Println("2. 格式化文本（粗体、斜体、下划线）")
	fmt.Println("3. 表格显示和边框")
	fmt.Println("4. 长文本换行")
	fmt.Println("5. 特殊字符显示")
	fmt.Println("6. 中英文混合显示")
	fmt.Println("7. 整体文档布局")
}
