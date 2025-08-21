package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("开始高级功能测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 测试1: 基本段落
	fmt.Println("1. 测试基本段落...")
	err = docWriter.AddParagraph("高级功能测试文档", "Normal")
	if err != nil {
		log.Fatal("Failed to add title paragraph:", err)
	}

	// 测试2: 复杂格式化段落
	fmt.Println("2. 测试复杂格式化段落...")
	complexRuns := []types.Run{
		{
			Text:     "标题：",
			FontName: "宋体",
			FontSize: 16,
			Bold:     true,
		},
		{
			Text:     "这是",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "重要",
			FontName: "宋体",
			FontSize: 14,
			Bold:     true,
		},
		{
			Text:     "的",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "信息",
			FontName: "宋体",
			FontSize: 14,
			Bold:     true,
		},
		{
			Text:     "。",
			FontName: "宋体",
			FontSize: 12,
		},
	}
	
	err = docWriter.AddFormattedParagraph("标题：这是重要的信息。", "Normal", complexRuns)
	if err != nil {
		log.Fatal("Failed to add complex formatted paragraph:", err)
	}

	// 测试3: 复杂表格
	fmt.Println("3. 测试复杂表格...")
	complexTableData := [][]string{
		{"项目", "负责人", "开始日期", "结束日期", "状态", "备注"},
		{"需求分析", "张三", "2024-01-01", "2024-01-15", "已完成", "按时完成"},
		{"系统设计", "李四", "2024-01-16", "2024-02-15", "进行中", "进度正常"},
		{"编码实现", "王五", "2024-02-16", "2024-04-15", "未开始", "等待设计完成"},
		{"测试验证", "赵六", "2024-04-16", "2024-05-15", "未开始", "等待编码完成"},
		{"部署上线", "钱七", "2024-05-16", "2024-06-01", "未开始", "等待测试完成"},
	}
	
	err = docWriter.AddTable(complexTableData)
	if err != nil {
		log.Fatal("Failed to add complex table:", err)
	}

	// 测试4: 多段落内容
	fmt.Println("4. 测试多段落内容...")
	paragraphs := []string{
		"项目概述：这是一个复杂的软件开发项目，涉及多个阶段和多个团队成员的协作。",
		"技术栈：项目使用现代化的技术栈，包括前端框架、后端服务、数据库和云服务等。",
		"团队结构：项目团队由项目经理、架构师、开发工程师、测试工程师和运维工程师组成。",
		"时间安排：整个项目预计需要6个月时间，分为5个主要阶段。",
		"风险控制：项目制定了详细的风险控制计划，包括技术风险、进度风险和人员风险等。",
	}

	for i, text := range paragraphs {
		err = docWriter.AddParagraph(text, "Normal")
		if err != nil {
			log.Fatal(fmt.Sprintf("Failed to add paragraph %d:", i+1), err)
		}
	}

	// 测试5: 特殊格式文本
	fmt.Println("5. 测试特殊格式文本...")
	specialRuns := []types.Run{
		{
			Text:     "代码示例：",
			FontName: "Consolas",
			FontSize: 12,
			Bold:     true,
		},
		{
			Text:     "func main() { fmt.Println(\"Hello, World!\") }",
			FontName: "Consolas",
			FontSize: 10,
		},
	}
	
	err = docWriter.AddFormattedParagraph("代码示例：func main() { fmt.Println(\"Hello, World!\") }", "Normal", specialRuns)
	if err != nil {
		log.Fatal("Failed to add code example:", err)
	}

	// 测试6: 数学公式和符号
	fmt.Println("6. 测试数学公式和符号...")
	mathText := "数学公式示例：E = mc²，π ≈ 3.14159，∑(i=1 to n) i = n(n+1)/2"
	err = docWriter.AddParagraph(mathText, "Normal")
	if err != nil {
		log.Fatal("Failed to add math text:", err)
	}

	// 测试7: 国际化文本
	fmt.Println("7. 测试国际化文本...")
	internationalText := "国际化测试：English 中文 日本語 한국어 Español Français Deutsch Italiano Português Русский العربية हिन्दी"
	err = docWriter.AddParagraph(internationalText, "Normal")
	if err != nil {
		log.Fatal("Failed to add international text:", err)
	}

	// 测试8: 长段落和换行
	fmt.Println("8. 测试长段落和换行...")
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

	// 测试9: 总结和评估
	fmt.Println("9. 添加总结和评估...")
	summaryRuns := []types.Run{
		{
			Text:     "测试总结：",
			FontName: "宋体",
			FontSize: 16,
			Bold:     true,
		},
		{
			Text:     "本次高级功能测试涵盖了复杂格式化、多列表格、长文本处理、特殊字符支持、国际化文本等高级功能。",
			FontName: "宋体",
			FontSize: 12,
		},
	}
	
	err = docWriter.AddFormattedParagraph("测试总结：本次高级功能测试涵盖了复杂格式化、多列表格、长文本处理、特殊字符支持、国际化文本等高级功能。", "Normal", summaryRuns)
	if err != nil {
		log.Fatal("Failed to add summary:", err)
	}

	// 保存文档
	filename := "advanced_feature_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n✅ 高级功能测试完成！文件已保存: %s\n", filename)
	fmt.Println("请打开这个文件，检查以下高级功能是否正确：")
	fmt.Println("1. 复杂格式化文本（多种字体大小、粗体组合）")
	fmt.Println("2. 多列表格（6列数据）")
	fmt.Println("3. 长段落自动换行")
	fmt.Println("4. 特殊字符和符号显示")
	fmt.Println("5. 国际化文本支持")
	fmt.Println("6. 代码示例格式")
	fmt.Println("7. 数学公式符号")
	fmt.Println("8. 整体文档布局和样式")
}
