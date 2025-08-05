package examples

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

func DemoAdvancedFeatures() {
	fmt.Println("=== Go Word 高级功能示例 ===\n")

	// 示例1：文档结构重组
	demoDocumentStructure()

	// 示例2：文档合并
	demoDocumentMerge()

	// 示例3：模板处理
	demoTemplateProcessing()

	fmt.Println("所有示例完成！")
}

// demoDocumentStructure 演示文档结构重组功能
func demoDocumentStructure() {
	fmt.Println("1. 文档结构重组示例")
	fmt.Println("-------------------")

	// 创建一个包含标题的文档
	w := writer.NewDocumentWriter()
	if err := w.CreateNewDocument(); err != nil {
		log.Printf("创建文档失败: %v", err)
		return
	}

	// 添加标题和内容
	w.AddParagraph("第一章 介绍", "Heading1")
	w.AddParagraph("这是第一章的内容。", "Normal")
	w.AddParagraph("第二章 方法", "Heading1")
	w.AddParagraph("这是第二章的内容。", "Normal")
	w.AddParagraph("2.1 子方法", "Heading2")
	w.AddParagraph("这是子方法的内容。", "Normal")

	// 保存文档
	if err := w.Save("structure_demo.docx"); err != nil {
		log.Printf("保存文档失败: %v", err)
		return
	}

	// 打开文档并分析结构
	doc, err := wordprocessingml.Open("structure_demo.docx")
	if err != nil {
		log.Printf("打开文档失败: %v", err)
		return
	}
	defer doc.Close()

	// 重组文档结构
	structure, err := doc.ReorganizeDocument()
	if err != nil {
		log.Printf("重组文档失败: %v", err)
		return
	}

	// 显示文档大纲
	fmt.Println("文档大纲:")
	fmt.Println(structure.GetOutlineAsString())

	// 获取所有段落
	sections := structure.GetSections()
	fmt.Printf("\n段落数量: %d\n", len(sections))

	// 按级别排序
	structure.SortSectionsByLevel()
	fmt.Println("按级别排序后的段落:")
	for _, section := range sections {
		fmt.Printf("- %s (级别: %d)\n", section.Title, section.Level)
	}

	fmt.Println()
}

// demoDocumentMerge 演示文档合并功能
func demoDocumentMerge() {
	fmt.Println("2. 文档合并示例")
	fmt.Println("----------------")

	// 创建第一个文档
	w1 := writer.NewDocumentWriter()
	w1.CreateNewDocument()
	w1.AddParagraph("第一个文档", "Heading1")
	w1.AddParagraph("这是第一个文档的内容。", "Normal")
	w1.Save("doc1.docx")

	// 创建第二个文档
	w2 := writer.NewDocumentWriter()
	w2.CreateNewDocument()
	w2.AddParagraph("第二个文档", "Heading1")
	w2.AddParagraph("这是第二个文档的内容。", "Normal")
	w2.Save("doc2.docx")

	// 打开目标文档
	targetDoc, err := wordprocessingml.Open("doc1.docx")
	if err != nil {
		log.Printf("打开目标文档失败: %v", err)
		return
	}
	defer targetDoc.Close()

	// 打开源文档
	sourceDoc, err := wordprocessingml.Open("doc2.docx")
	if err != nil {
		log.Printf("打开源文档失败: %v", err)
		return
	}
	defer sourceDoc.Close()

	// 创建合并操作
	merge := wordprocessingml.NewDocumentMerge(targetDoc)
	merge.AddSourceDocument(sourceDoc)

	// 设置合并选项
	options := wordprocessingml.MergeOptions{
		MergeMode:          wordprocessingml.AppendMode,
		ConflictResolution: wordprocessingml.KeepFirst,
		PreserveFormatting: true,
		AddPageBreaks:      true,
		IncludeTables:      true,
		IncludeImages:      true,
		MergeStyles:        true,
	}
	merge.SetMergeOptions(options)

	// 验证合并操作
	if err := merge.ValidateMerge(); err != nil {
		log.Printf("合并验证失败: %v", err)
		return
	}

	// 执行合并
	if err := merge.MergeDocuments(); err != nil {
		log.Printf("合并文档失败: %v", err)
		return
	}

	// 显示合并摘要
	fmt.Println("合并摘要:")
	fmt.Println(merge.GetMergeSummary())

	// 保存合并后的文档
	container := targetDoc.GetContainer()
	if err := container.SaveToFile("merged_document.docx"); err != nil {
		log.Printf("保存合并文档失败: %v", err)
		return
	}

	fmt.Println("文档合并完成！")
	fmt.Println()
}

// demoTemplateProcessing 演示模板处理功能
func demoTemplateProcessing() {
	fmt.Println("3. 模板处理示例")
	fmt.Println("----------------")

	// 创建一个模板文档
	w := writer.NewDocumentWriter()
	w.CreateNewDocument()
	w.AddParagraph("{{title}}", "Heading1")
	w.AddParagraph("作者: {{author}}", "Normal")
	w.AddParagraph("日期: {{date}}", "Normal")
	w.AddParagraph("价格: {{price}} 元", "Normal")
	w.AddParagraph("是否可用: {{available}}", "Normal")
	w.AddParagraph("表格数据:", "Normal")
	w.AddParagraph("{{table_data}}", "Normal")
	w.Save("template.docx")

	// 打开模板文档
	doc, err := wordprocessingml.Open("template.docx")
	if err != nil {
		log.Printf("打开模板文档失败: %v", err)
		return
	}
	defer doc.Close()

	// 创建模板处理器
	template := wordprocessingml.NewTemplate(doc)

	// 提取占位符
	if err := template.ExtractPlaceholders(); err != nil {
		log.Printf("提取占位符失败: %v", err)
		return
	}

	// 显示模板摘要
	fmt.Println("模板摘要:")
	fmt.Println(template.GetTemplateSummary())

	// 添加变量
	template.AddVariable("title", "示例文档")
	template.AddVariable("author", "张三")
	template.AddVariable("date", "2024-01-01")
	template.AddVariable("price", 99.99)
	template.AddVariable("available", true)
	template.AddVariable("table_data", [][]string{
		{"产品", "价格", "库存"},
		{"产品A", "100", "10"},
		{"产品B", "200", "5"},
	})

	// 添加自定义占位符
	placeholder := wordprocessingml.TemplatePlaceholder{
		Key:         "title",
		Type:        wordprocessingml.TextPlaceholder,
		Required:    true,
		DefaultValue: "默认标题",
	}
	template.AddPlaceholder(placeholder)

	// 处理模板
	if err := template.ProcessTemplate(); err != nil {
		log.Printf("处理模板失败: %v", err)
		return
	}

	// 保存处理后的文档
	container := doc.GetContainer()
	if err := container.SaveToFile("processed_template.docx"); err != nil {
		log.Printf("保存处理后的模板失败: %v", err)
		return
	}

	fmt.Println("模板处理完成！")
	fmt.Println()
} 