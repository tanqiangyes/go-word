package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/writer"
)

func main() {
	fmt.Println("🔧 开始修复后的批注功能测试...")

	// 创建文档写入器
	docWriter := writer.NewDocumentWriter()

	// 创建新文档
	err := docWriter.CreateNewDocument()
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}

	// 测试1: 添加标题
	fmt.Println("1. 添加文档标题...")
	err = docWriter.AddParagraph("修复后的批注功能测试文档", "Normal")
	if err != nil {
		log.Fatal("Failed to add title:", err)
	}

	// 测试2: 添加段落并添加批注
	fmt.Println("2. 添加段落并添加批注...")
	
	// 第一个段落
	err = docWriter.AddParagraph("这是第一个段落，包含重要信息，需要高亮显示。", "Normal")
	if err != nil {
		log.Fatal("Failed to add first paragraph:", err)
	}
	
	// 为第一个段落添加批注
	err = docWriter.AddComment("张三", "这个段落包含重要信息，需要高亮显示。", "这是第一个段落，包含重要信息，需要高亮显示。")
	if err != nil {
		log.Fatal("Failed to add comment to first paragraph:", err)
	}
	fmt.Println("✅ 为第一个段落添加批注成功")

	// 第二个段落
	err = docWriter.AddParagraph("这是第二个段落，需要审查和修改。", "Normal")
	if err != nil {
		log.Fatal("Failed to add second paragraph:", err)
	}
	
	// 为第二个段落添加批注
	err = docWriter.AddComment("李四", "建议在这里添加更多详细信息，使内容更加完整。", "这是第二个段落，需要审查和修改。")
	if err != nil {
		log.Fatal("Failed to add comment to second paragraph:", err)
	}
	fmt.Println("✅ 为第二个段落添加批注成功")

	// 第三个段落
	err = docWriter.AddParagraph("这是第三个段落，内容很好，可以保留。", "Normal")
	if err != nil {
		log.Fatal("Failed to add third paragraph:", err)
	}
	
	// 为第三个段落添加批注
	err = docWriter.AddComment("王五", "这个段落内容很好，可以保留。建议在其他地方也采用类似的写作风格。", "这是第三个段落，内容很好，可以保留。")
	if err != nil {
		log.Fatal("Failed to add comment to third paragraph:", err)
	}
	fmt.Println("✅ 为第三个段落添加批注成功")

	// 测试3: 添加表格
	fmt.Println("3. 添加表格...")
	tableData := [][]string{
		{"项目", "状态", "负责人", "备注"},
		{"需求分析", "已完成", "张三", "需要进一步细化"},
		{"系统设计", "进行中", "李四", "设计文档待完善"},
		{"编码实现", "未开始", "王五", "等待设计完成"},
	}
	
	err = docWriter.AddTable(tableData)
	if err != nil {
		log.Fatal("Failed to add table:", err)
	}
	fmt.Println("✅ 表格添加成功")

	// 测试4: 添加长段落并添加批注
	fmt.Println("4. 添加长段落并添加批注...")
	longParagraph := "这是一个非常长的段落，用来测试批注功能在长文本中的表现。" +
		"当文本内容很长时，批注应该能够正确关联到对应的段落。" +
		"长文本段落可以帮助我们验证批注功能的稳定性和准确性。" +
		"这个测试段落包含了多个句子，涵盖了不同的标点符号和格式要求。" +
		"我们还需要测试批注在不同字体大小下的显示效果，以及在不同页面宽度下的换行行为。"
	
	err = docWriter.AddParagraph(longParagraph, "Normal")
	if err != nil {
		log.Fatal("Failed to add long paragraph:", err)
	}
	
	// 为长段落添加批注
	err = docWriter.AddComment("编辑", "这个长段落内容很好，但建议分成几个小段落，提高可读性。", longParagraph)
	if err != nil {
		log.Fatal("Failed to add comment to long paragraph:", err)
	}
	fmt.Println("✅ 为长段落添加批注成功")

	// 测试5: 添加总结段落
	fmt.Println("5. 添加总结段落...")
	err = docWriter.AddParagraph("修复后的批注功能测试总结：本次测试验证了修复后的批注功能，应该能够正常显示批注内容。", "Normal")
	if err != nil {
		log.Fatal("Failed to add summary paragraph:", err)
	}

	// 保存文档
	filename := "fixed_comment_test.docx"
	err = docWriter.Save(filename)
	if err != nil {
		log.Fatal("Failed to save document:", err)
	}

	fmt.Printf("\n🎉 修复后的批注功能测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容概览：")
	fmt.Println("1. 文档标题")
	fmt.Println("2. 第一个段落 + 张三的批注")
	fmt.Println("3. 第二个段落 + 李四的批注")
	fmt.Println("4. 第三个段落 + 王五的批注")
	fmt.Println("5. 项目进度表格（无批注）")
	fmt.Println("6. 长段落 + 编辑的批注")
	fmt.Println("7. 总结段落")
	
	fmt.Println("\n🔧 修复内容：")
	fmt.Println("- 修复了批注 XML 的命名空间问题")
	fmt.Println("- 修复了设置文件的命名空间问题")
	fmt.Println("- 使用正确的 wordprocessingml 命名空间")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 批注是否正确显示在文档中")
	fmt.Println("- 批注作者信息是否正确")
	fmt.Println("- 批注内容是否完整")
	fmt.Println("- 批注是否与正确的段落关联")
	fmt.Println("- 在 Word 中是否能正常查看批注")
	fmt.Println("- 批注的显示和隐藏是否正常")
	
	fmt.Println("\n💡 查看批注的方法：")
	fmt.Println("1. 在 Word 中打开文档")
	fmt.Println("2. 点击 '审阅' 选项卡")
	fmt.Println("3. 点击 '显示批注' 按钮")
	fmt.Println("4. 批注应该显示在右侧边栏中")
	
	fmt.Println("\n🏆 如果批注现在能正常显示，说明问题已经修复！")
}
