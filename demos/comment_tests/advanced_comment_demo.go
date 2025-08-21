package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🚀 开始高级批注功能测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 测试1: 文档标题
	fmt.Println("1. 添加文档标题...")
	err = docWriter.AddParagraph("高级批注功能测试文档", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	// 测试2: 项目概述段落
	fmt.Println("2. 添加项目概述段落...")
	overviewText := "项目概述：这是一个复杂的软件开发项目，涉及多个阶段和多个团队成员的协作。项目目标是开发一个现代化的企业级应用系统。"
	err = docWriter.AddParagraph(overviewText, "Normal")
	if err != nil {
		log.Fatal("Failed to add overview paragraph:", err)
	}
	
	// 为项目概述添加批注
	err = docWriter.AddComment("项目经理", "项目概述写得很好，但建议添加具体的项目时间线和里程碑信息。", overviewText)
	if err != nil {
		log.Fatal("Failed to add comment to overview:", err)
	}
	fmt.Println("✅ 为项目概述添加批注成功")

	// 测试3: 技术栈段落
	fmt.Println("3. 添加技术栈段落...")
	techStackText := "技术栈：项目使用现代化的技术栈，包括前端框架（React/Vue.js）、后端服务（Go/Node.js）、数据库（PostgreSQL/MongoDB）和云服务（AWS/Azure）等。"
	err = docWriter.AddParagraph(techStackText, "Normal")
	if err != nil {
		log.Fatal("Failed to add tech stack paragraph:", err)
	}
	
	// 为技术栈添加批注
	err = docWriter.AddComment("架构师", "技术选型合理，但建议考虑微服务架构的复杂性，需要评估团队的技术能力。", techStackText)
	if err != nil {
		log.Fatal("Failed to add comment to tech stack:", err)
	}
	fmt.Println("✅ 为技术栈添加批注成功")

	// 测试4: 团队结构段落
	fmt.Println("4. 添加团队结构段落...")
	teamText := "团队结构：项目团队由项目经理、架构师、开发工程师、测试工程师和运维工程师组成，总计15人。"
	err = docWriter.AddParagraph(teamText, "Normal")
	if err != nil {
		log.Fatal("Failed to add team structure paragraph:", err)
	}
	
	// 为团队结构添加批注
	err = docWriter.AddComment("HR经理", "团队规模合适，但建议增加一名产品经理和一名UI/UX设计师，以提升产品体验。", teamText)
	if err != nil {
		log.Fatal("Failed to add comment to team structure:", err)
	}
	fmt.Println("✅ 为团队结构添加批注成功")

	// 测试5: 格式化文本段落
	fmt.Println("5. 添加格式化文本段落...")
	formattedRuns := []types.Run{
		{
			Text:     "重要提示：",
			FontName: "宋体",
			FontSize: 14,
			Bold:     true,
		},
		{
			Text:     "这个项目需要特别注意",
			FontName: "宋体",
			FontSize: 12,
		},
		{
			Text:     "安全性和性能",
			FontName: "宋体",
			FontSize: 12,
			Bold:     true,
		},
		{
			Text:     "方面的要求。",
			FontName: "宋体",
			FontSize: 12,
		},
	}
	
	err = docWriter.AddFormattedParagraph("重要提示：这个项目需要特别注意安全性和性能方面的要求。", "Normal", formattedRuns)
	if err != nil {
		log.Fatal("Failed to add formatted paragraph:", err)
	}
	
	// 为格式化段落添加批注
	err = docWriter.AddComment("安全专家", "安全要求非常重要，建议在项目初期就制定详细的安全规范和测试计划。", "重要提示：这个项目需要特别注意安全性和性能方面的要求。")
	if err != nil {
		log.Fatal("Failed to add comment to formatted paragraph:", err)
	}
	fmt.Println("✅ 为格式化段落添加批注成功")

	// 测试6: 项目时间表
	fmt.Println("6. 添加项目时间表...")
	tableData := [][]string{
		{"阶段", "开始时间", "结束时间", "负责人", "状态", "备注"},
		{"需求分析", "2024-01-01", "2024-01-31", "张三", "已完成", "需求文档已确认"},
		{"系统设计", "2024-02-01", "2024-03-31", "李四", "进行中", "架构设计待评审"},
		{"编码实现", "2024-04-01", "2024-08-31", "王五", "未开始", "等待设计完成"},
		{"测试验证", "2024-09-01", "2024-10-31", "赵六", "未开始", "等待编码完成"},
		{"部署上线", "2024-11-01", "2024-12-31", "钱七", "未开始", "等待测试完成"},
	}
	
	err = docWriter.AddTable(tableData)
	if err != nil {
		log.Fatal("Failed to add project timeline table:", err)
	}
	fmt.Println("✅ 项目时间表添加成功")

	// 测试7: 风险分析段落
	fmt.Println("7. 添加风险分析段落...")
	riskText := "风险分析：项目面临的主要风险包括技术风险（新技术学习成本）、进度风险（依赖外部系统）、人员风险（关键人员流失）和预算风险（成本超支）。"
	err = docWriter.AddParagraph(riskText, "Normal")
	if err != nil {
		log.Fatal("Failed to add risk analysis paragraph:", err)
	}
	
	// 为风险分析添加批注
	err = docWriter.AddComment("风险经理", "风险识别全面，建议制定详细的风险应对策略和应急预案，并定期进行风险评估。", riskText)
	if err != nil {
		log.Fatal("Failed to add comment to risk analysis:", err)
	}
	fmt.Println("✅ 为风险分析添加批注成功")

	// 测试8: 质量保证段落
	fmt.Println("8. 添加质量保证段落...")
	qualityText := "质量保证：项目将采用敏捷开发方法，每个迭代都进行代码审查、单元测试和集成测试。同时建立持续集成/持续部署（CI/CD）流程，确保代码质量。"
	err = docWriter.AddParagraph(qualityText, "Normal")
	if err != nil {
		log.Fatal("Failed to add quality assurance paragraph:", err)
	}
	
	// 为质量保证添加批注
	err = docWriter.AddComment("质量经理", "质量保证策略很好，建议增加自动化测试覆盖率要求，并建立代码质量度量指标。", qualityText)
	if err != nil {
		log.Fatal("Failed to add comment to quality assurance:", err)
	}
	fmt.Println("✅ 为质量保证添加批注成功")

	// 测试9: 总结段落
	fmt.Println("9. 添加总结段落...")
	summaryText := "项目总结：这是一个具有挑战性的企业级应用开发项目，需要团队成员的密切协作和持续改进。通过合理的规划和管理，我们有信心按时交付高质量的软件产品。"
	err = docWriter.AddParagraph(summaryText, "Normal")
	if err != nil {
		log.Fatal("Failed to add summary paragraph:", err)
	}
	
	// 为总结添加批注
	err = docWriter.AddComment("项目总监", "总结写得很好，体现了团队的专业性和信心。建议在项目执行过程中定期回顾和调整计划。", summaryText)
	if err != nil {
		log.Fatal("Failed to add comment to summary:", err)
	}
	fmt.Println("✅ 为总结添加批注成功")

	// 保存文档
	filename := "advanced_comment_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n🎉 高级批注功能测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容概览：")
	fmt.Println("1. 文档标题")
	fmt.Println("2. 项目概述 + 项目经理批注")
	fmt.Println("3. 技术栈 + 架构师批注")
	fmt.Println("4. 团队结构 + HR经理批注")
	fmt.Println("5. 格式化文本 + 安全专家批注")
	fmt.Println("6. 项目时间表（无批注）")
	fmt.Println("7. 风险分析 + 风险经理批注")
	fmt.Println("8. 质量保证 + 质量经理批注")
	fmt.Println("9. 项目总结 + 项目总监批注")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 所有批注是否正确显示")
	fmt.Println("- 批注作者信息是否准确")
	fmt.Println("- 批注内容是否完整")
	fmt.Println("- 批注是否与正确段落关联")
	fmt.Println("- 格式化文本的批注是否正常")
	fmt.Println("- 长文本的批注是否稳定")
	
	fmt.Println("\n💡 查看批注的方法：")
	fmt.Println("1. 在 Word 中打开文档")
	fmt.Println("2. 点击 '审阅' 选项卡")
	fmt.Println("3. 点击 '显示批注' 按钮")
	fmt.Println("4. 批注应该显示在右侧边栏中")
	fmt.Println("5. 点击批注可以跳转到对应段落")
	
	fmt.Println("\n🏆 如果所有批注都能正常显示，说明高级批注功能已经完善！")
	fmt.Println("🚀 可以用于复杂的项目文档和协作场景！")
}
