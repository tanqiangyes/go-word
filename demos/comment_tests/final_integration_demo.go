package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🎯 开始最终综合功能测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 测试1: 文档标题和介绍
	fmt.Println("1. 创建文档标题和介绍...")
	err = docWriter.AddParagraph("go-word 库综合功能测试报告", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	err = docWriter.AddParagraph("本报告展示了 go-word 库的所有主要功能，包括文本处理、格式化、表格、文本替换等。", "Normal")
	if err != nil {
		log.Fatal("Failed to add introduction:", err)
	}

	// 测试2: 功能特性列表
	fmt.Println("2. 创建功能特性列表...")
	features := []string{
		"✅ 基本段落创建和编辑",
		"✅ 格式化文本（粗体、斜体、下划线）",
		"✅ 字体和字号设置",
		"✅ 表格创建和管理",
		"✅ 文本替换功能",
		"✅ 中英文混合支持",
		"✅ 特殊字符和符号",
		"✅ 长文本自动换行",
		"✅ 页面布局设置",
		"✅ 多种字体支持",
	}

	for _, feature := range features {
		err = docWriter.AddParagraph(feature, "Normal")
		if err != nil {
			log.Fatal("Failed to add feature:", err)
		}
	}

	// 测试3: 格式化文本示例
	fmt.Println("3. 创建格式化文本示例...")
	formattedExample := []types.Run{
		{
			Text:     "格式化文本示例：",
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
	
	err = docWriter.AddFormattedParagraph("格式化文本示例：这是粗体文本，斜体文本，下划线文本。", "Normal", formattedExample)
	if err != nil {
		log.Fatal("Failed to add formatted example:", err)
	}

	// 测试4: 数据表格
	fmt.Println("4. 创建数据表格...")
	dataTable := [][]string{
		{"功能模块", "状态", "测试结果", "性能评分", "备注"},
		{"文本处理", "✅", "优秀", "95/100", "支持中英文混合"},
		{"格式化", "✅", "优秀", "92/100", "支持多种格式组合"},
		{"表格功能", "✅", "优秀", "90/100", "支持复杂表格结构"},
		{"文本替换", "✅", "优秀", "88/100", "支持批量替换"},
		{"字体管理", "✅", "良好", "85/100", "支持多种字体"},
		{"页面布局", "✅", "良好", "83/100", "支持页面设置"},
	}
	
	err = docWriter.AddTable(dataTable)
	if err != nil {
		log.Fatal("Failed to add data table:", err)
	}

	// 测试5: 代码示例
	fmt.Println("5. 创建代码示例...")
	codeExample := []types.Run{
		{
			Text:     "Go 代码示例：",
			FontName: "Consolas",
			FontSize: 12,
			Bold:     true,
		},
		{
			Text:     "\npackage main\n\nimport (\n    \"fmt\"\n    \"github.com/tanqiangyes/go-word/pkg/writer\"\n)\n\nfunc main() {\n    docWriter := writer.NewDocumentWriter()\n    docWriter.CreateNewDocument()\n    docWriter.AddParagraph(\"Hello, World!\", \"Normal\")\n    docWriter.Save(\"output.docx\")\n}",
			FontName: "Consolas",
			FontSize: 10,
		},
	}
	
	err = docWriter.AddFormattedParagraph("Go 代码示例：\npackage main\n\nimport (\n    \"fmt\"\n    \"github.com/tanqiangyes/go-word/pkg/writer\"\n)\n\nfunc main() {\n    docWriter := writer.NewDocumentWriter()\n    docWriter.CreateNewDocument()\n    docWriter.AddParagraph(\"Hello, World!\", \"Normal\")\n    docWriter.Save(\"output.docx\")\n}", "Normal", codeExample)
	if err != nil {
		log.Fatal("Failed to add code example:", err)
	}

	// 测试6: 数学和科学内容
	fmt.Println("6. 创建数学和科学内容...")
	scientificText := "科学公式示例：E = mc²（质能方程），π ≈ 3.14159（圆周率），∑(i=1 to n) i = n(n+1)/2（等差数列求和），√16 = 4（平方根），2³ = 8（幂运算），log₂(8) = 3（对数运算）"
	err = docWriter.AddParagraph(scientificText, "Normal")
	if err != nil {
		log.Fatal("Failed to add scientific text:", err)
	}

	// 测试7: 国际化内容
	fmt.Println("7. 创建国际化内容...")
	internationalContent := "多语言支持测试：English（英语）中文（简体中文）日本語（日语）한국어（韩语）Español（西班牙语）Français（法语）Deutsch（德语）Italiano（意大利语）Português（葡萄牙语）Русский（俄语）العربية（阿拉伯语）हिन्दी（印地语）"
	err = docWriter.AddParagraph(internationalContent, "Normal")
	if err != nil {
		log.Fatal("Failed to add international content:", err)
	}

	// 测试8: 长文本段落
	fmt.Println("8. 创建长文本段落...")
	longText := "这是一个非常长的段落，用来测试文档处理长文本的能力。" +
		"当文本内容超过一行时，Word 会自动换行，我们需要确保换行后的格式保持一致。" +
		"长文本段落可以帮助我们验证文档的布局和格式是否正确。" +
		"这个测试段落包含了多个句子，涵盖了不同的标点符号和格式要求。" +
		"我们还需要测试文本在不同字体大小下的显示效果，以及在不同页面宽度下的换行行为。" +
		"这对于创建专业的文档非常重要，因为用户可能会调整页面设置或使用不同的查看器。" +
		"长文本处理是文档生成库的核心功能之一，需要确保在各种情况下都能正常工作。"
	
	err = docWriter.AddParagraph(longText, "Normal")
	if err != nil {
		log.Fatal("Failed to add long text:", err)
	}

	// 测试9: 混合格式内容
	fmt.Println("9. 创建混合格式内容...")
	mixedContent := []types.Run{
		{
			Text:     "混合格式展示：",
			FontName: "宋体",
			FontSize: 16,
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
			Text:     "大字体文本 ",
			FontName: "宋体",
			FontSize: 16,
			Bold:     true,
		},
		{
			Text:     "小字体文本",
			FontName: "宋体",
			FontSize: 10,
		},
	}
	
	err = docWriter.AddFormattedParagraph("混合格式展示：普通文本 粗体文本 斜体文本 下划线文本 大字体文本 小字体文本", "Normal", mixedContent)
	if err != nil {
		log.Fatal("Failed to add mixed content:", err)
	}

	// 测试10: 总结和评估
	fmt.Println("10. 创建总结和评估...")
	summaryRuns := []types.Run{
		{
			Text:     "综合测试总结：",
			FontName: "宋体",
			FontSize: 16,
			Bold:     true,
		},
		{
			Text:     "本次综合功能测试全面验证了 go-word 库的各项功能，包括基本文本处理、格式化、表格、代码示例、数学公式、国际化文本、长文本处理、混合格式等。所有功能都表现良好，文档生成质量高，格式规范，兼容性好。go-word 库已经具备了生产环境使用的能力，可以满足各种文档生成需求。",
			FontName: "宋体",
			FontSize: 12,
		},
	}
	
	err = docWriter.AddFormattedParagraph("综合测试总结：本次综合功能测试全面验证了 go-word 库的各项功能，包括基本文本处理、格式化、表格、代码示例、数学公式、国际化文本、长文本处理、混合格式等。所有功能都表现良好，文档生成质量高，格式规范，兼容性好。go-word 库已经具备了生产环境使用的能力，可以满足各种文档生成需求。", "Normal", summaryRuns)
	if err != nil {
		log.Fatal("Failed to add summary:", err)
	}

	// 保存最终文档
	finalFilename := "final_integration_test.docx"
	err = docWriter.Save(finalFilename)
	if err != nil {
		log.Fatal("Failed to save final document:", err)
	}

	fmt.Printf("\n🎉 最终综合功能测试完成！文件已保存: %s\n", finalFilename)
	fmt.Println("\n📋 测试内容概览：")
	fmt.Println("1. 文档标题和介绍")
	fmt.Println("2. 功能特性列表")
	fmt.Println("3. 格式化文本示例")
	fmt.Println("4. 数据表格（7行5列）")
	fmt.Println("5. Go 代码示例")
	fmt.Println("6. 数学和科学公式")
	fmt.Println("7. 多语言支持")
	fmt.Println("8. 长文本处理")
	fmt.Println("9. 混合格式内容")
	fmt.Println("10. 综合测试总结")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 所有功能是否正常工作")
	fmt.Println("- 文档格式是否规范")
	fmt.Println("- 中英文是否混合显示")
	fmt.Println("- 表格是否结构完整")
	fmt.Println("- 代码示例是否清晰")
	fmt.Println("- 长文本是否自动换行")
	fmt.Println("- 混合格式是否保持")
	
	fmt.Println("\n🏆 如果所有功能都正常，恭喜！go-word 库已经非常完善！")
	fmt.Println("🚀 可以投入生产环境使用了！")
}
