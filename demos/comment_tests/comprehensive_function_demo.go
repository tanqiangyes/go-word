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

	// 测试1: 标题和基本段落
	fmt.Println("1. 测试标题和基本段落...")
	err = docWriter.AddParagraph("全面功能测试文档", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	err = docWriter.AddParagraph("这是一个全面的功能测试文档，将测试 go-word 库的所有主要功能。", "Normal")
	if err != nil {
		log.Fatal("Failed to add description:", err)
	}

	// 测试2: 格式化文本
	fmt.Println("2. 测试格式化文本...")
	formattedRuns := []types.Run{
		{
			Text:     "重要提示：",
			FontName: "宋体",
			FontSize: 14,
			Bold:     true,
		},
		{
			Text:     "这是",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "粗体",
			FontName: "宋体",
			FontSize: 12,
			Bold:     true,
		},
		{
			Text:     "文本，",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "斜体",
			FontName: "宋体",
			FontSize: 12,
			Italic:   true,
		},
		{
			Text:     "文本，",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "下划线",
			FontName: "宋体",
			FontSize: 12,
			Underline: true,
		},
		{
			Text:     "文本。",
			FontName: "宋体",
			FontSize: 12,
		},
	}
	
	err = docWriter.AddFormattedParagraph("重要提示：这是粗体文本，斜体文本，下划线文本。", "Normal", formattedRuns)
	if err != nil {
		log.Fatal("Failed to add formatted text:", err)
	}

	// 测试3: 复杂表格
	fmt.Println("3. 测试复杂表格...")
	complexTableData := [][]string{
		{"功能模块", "状态", "测试结果", "备注", "优先级"},
		{"基本段落", "✅", "通过", "支持中英文", "高"},
		{"格式化文本", "✅", "通过", "支持粗体、斜体、下划线", "高"},
		{"表格功能", "✅", "通过", "支持多行多列", "高"},
		{"文本替换", "✅", "通过", "支持占位符替换", "中"},
		{"字体设置", "✅", "通过", "支持多种字体", "中"},
		{"页面设置", "✅", "通过", "支持页面大小和边距", "低"},
	}
	
	err = docWriter.AddTable(complexTableData)
	if err != nil {
		log.Fatal("Failed to add complex table:", err)
	}

	// 测试4: 代码示例
	fmt.Println("4. 测试代码示例...")
	codeRuns := []types.Run{
		{
			Text:     "Go 代码示例：",
			FontName: "Consolas",
			FontSize: 12,
			Bold:     true,
		},
		{
			Text:     "\npackage main\n\nimport \"fmt\"\n\nfunc main() {\n    fmt.Println(\"Hello, World!\")\n}",
			FontName: "Consolas",
			FontSize: 10,
		},
	}
	
	err = docWriter.AddFormattedParagraph("Go 代码示例：\npackage main\n\nimport \"fmt\"\n\nfunc main() {\n    fmt.Println(\"Hello, World!\")\n}", "Normal", codeRuns)
	if err != nil {
		log.Fatal("Failed to add code example:", err)
	}

	// 测试5: 数学公式和特殊符号
	fmt.Println("5. 测试数学公式和特殊符号...")
	mathText := "数学公式示例：E = mc²，π ≈ 3.14159，∑(i=1 to n) i = n(n+1)/2，√16 = 4，2³ = 8"
	err = docWriter.AddParagraph(mathText, "Normal")
	if err != nil {
		log.Fatal("Failed to add math text:", err)
	}

	// 测试6: 国际化文本
	fmt.Println("6. 测试国际化文本...")
	internationalText := "国际化测试：English 中文 日本語 한국어 Español Français Deutsch Italiano Português Русский العربية हिन्दी"
	err = docWriter.AddParagraph(internationalText, "Normal")
	if err != nil {
		log.Fatal("Failed to add international text:", err)
	}

	// 测试7: 长段落和换行
	fmt.Println("7. 测试长段落和换行...")
	longParagraph := "这是一个非常长的段落，用来测试文档处理长文本的能力。" +
		"当文本内容超过一行时，Word 会自动换行，我们需要确保换行后的格式保持一致。" +
		"长文本段落可以帮助我们验证文档的布局和格式是否正确。" +
		"这个测试段落包含了多个句子，涵盖了不同的标点符号和格式要求。" +
		"我们还需要测试文本在不同字体大小下的显示效果，以及在不同页面宽度下的换行行为。" +
		"这对于创建专业的文档非常重要，因为用户可能会调整页面设置或使用不同的查看器。"
	
	err = docWriter.AddParagraph(longParagraph, "Normal")
	if err != nil {
		log.Fatal("Failed to add long paragraph:", err)
	}

	// 测试8: 列表和编号
	fmt.Println("8. 测试列表和编号...")
	listItems := []string{
		"第一项：基本功能测试",
		"第二项：格式化功能测试", 
		"第三项：表格功能测试",
		"第四项：文本替换功能测试",
		"第五项：字体和样式测试",
		"第六项：页面布局测试",
	}

	for i, item := range listItems {
		err = docWriter.AddParagraph(fmt.Sprintf("%d. %s", i+1, item), "Normal")
		if err != nil {
			log.Fatal(fmt.Sprintf("Failed to add list item %d:", i+1), err)
		}
	}

	// 测试9: 混合内容
	fmt.Println("9. 测试混合内容...")
	mixedRuns := []types.Run{
		{
			Text:     "混合格式：",
			FontName: "宋体",
			FontSize: 14,
			Bold:     true,
		},
		{
			Text:     "普通文本 ",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "粗体文本 ",
			FontName: "宋体",
			FontSize: 12,
			Bold:     true,
		},
		{
			Text:     "斜体文本 ",
			FontName: "宋体",
			FontSize: 12,
			Italic:   true,
		},
		{
			Text:     "下划线文本 ",
			FontName: "宋体",
			FontSize: 12,
			Underline: true,
		},
		{
			Text:     "大字体文本",
			FontName: "宋体",
			FontSize: 16,
			Bold:     true,
		},
	}
	
	err = docWriter.AddFormattedParagraph("混合格式：普通文本 粗体文本 斜体文本 下划线文本 大字体文本", "Normal", mixedRuns)
	if err != nil {
		log.Fatal("Failed to add mixed format text:", err)
	}

	// 测试10: 总结
	fmt.Println("10. 添加测试总结...")
	summaryRuns := []types.Run{
		{
			Text:     "测试总结：",
			FontName: "宋体",
			FontSize: 16,
			Bold:     true,
		},
		{
			Text:     "本次全面功能测试涵盖了 go-word 库的所有主要功能，包括基本段落创建、格式化文本、表格、代码示例、数学公式、国际化文本、长文本处理、列表编号、混合格式等。如果这个文档能够正常打开和显示所有内容，说明我们的 DOCX 生成功能已经非常完善。",
			FontName: "宋体",
			FontSize: 12,
		},
	}
	
	err = docWriter.AddFormattedParagraph("测试总结：本次全面功能测试涵盖了 go-word 库的所有主要功能，包括基本段落创建、格式化文本、表格、代码示例、数学公式、国际化文本、长文本处理、列表编号、混合格式等。如果这个文档能够正常打开和显示所有内容，说明我们的 DOCX 生成功能已经非常完善。", "Normal", summaryRuns)
	if err != nil {
		log.Fatal("Failed to add summary:", err)
	}

	// 保存文档
	filename := "comprehensive_function_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n✅ 全面功能测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n请打开这个文件，检查以下功能是否正确：")
	fmt.Println("1. 基本段落和标题")
	fmt.Println("2. 格式化文本（粗体、斜体、下划线、不同字体大小）")
	fmt.Println("3. 复杂表格（7行5列数据）")
	fmt.Println("4. 代码示例（等宽字体）")
	fmt.Println("5. 数学公式和特殊符号")
	fmt.Println("6. 国际化文本支持")
	fmt.Println("7. 长段落自动换行")
	fmt.Println("8. 列表和编号")
	fmt.Println("9. 混合格式文本")
	fmt.Println("10. 整体文档布局和样式")
	fmt.Println("\n如果所有功能都正常，说明 go-word 库已经非常完善！")
}
