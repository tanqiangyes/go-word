package examples

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

// 基本使用示例
// 演示如何打开Word文档并提取其内容
func DemoBasicUsage() {
	fmt.Println("=== Go OpenXML SDK 基本使用示例 ===")
	fmt.Println("本示例演示如何打开Word文档并提取文本、段落和表格信息")
	fmt.Println()

	// 使用示例文档
	filename := "example.docx"

	fmt.Printf("正在打开文档: %s\n", filename)

	// 打开Word文档
	// Open函数返回一个Document对象，该对象包含文档的所有内容
	doc, err := wordprocessingml.Open(filename)
	if err != nil {
		log.Printf("❌ 无法打开文档 %s: %v", filename, err)
		fmt.Println("💡 提示:")
		fmt.Println("   - 请确保文件路径正确")
		fmt.Println("   - 请确保文件是有效的.docx格式")
		fmt.Println("   - 请确保文件没有被其他程序占用")
		fmt.Println()
		fmt.Println("📝 使用方法:")
		fmt.Printf("   调用 DemoBasicUsage() 函数\n")
		return
	}
	defer doc.Close() // 确保文档资源被释放

	fmt.Println("✅ 文档打开成功")
	fmt.Println()

	// 示例1: 获取文档的纯文本内容
	fmt.Println("📄 示例1: 提取文档文本")
	text, err := doc.GetText()
	if err != nil {
		log.Printf("❌ 无法获取文档文本: %v", err)
		return
	}

	if text == "" {
		fmt.Println("⚠️  文档中没有文本内容")
	} else {
		fmt.Printf("📝 文档文本内容 (%d 字符):\n", len(text))
		// 只显示前200个字符，避免输出过长
		if len(text) > 200 {
			fmt.Printf("%s...\n", text[:200])
		} else {
			fmt.Printf("%s\n", text)
		}
	}
	fmt.Println()

	// 示例2: 获取文档中的所有段落
	fmt.Println("📄 示例2: 提取段落信息")
	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		log.Printf("❌ 无法获取段落: %v", err)
		return
	}

	fmt.Printf("📊 文档包含 %d 个段落\n", len(paragraphs))
	for i, paragraph := range paragraphs {
		fmt.Printf("   段落 %d: ", i+1)
		if paragraph.Text == "" {
			fmt.Println("(空段落)")
		} else {
			// 只显示前50个字符
			if len(paragraph.Text) > 50 {
				fmt.Printf("%s...\n", paragraph.Text[:50])
			} else {
				fmt.Printf("%s\n", paragraph.Text)
			}
		}

		// 显示段落的格式化信息
		if len(paragraph.Runs) > 0 {
			fmt.Printf("     包含 %d 个文本运行\n", len(paragraph.Runs))
			for j, run := range paragraph.Runs {
				fmt.Printf("       运行 %d: '%s'", j+1, run.Text)
				if run.Bold {
					fmt.Print(" [粗体]")
				}
				if run.Italic {
					fmt.Print(" [斜体]")
				}
				if run.Underline {
					fmt.Print(" [下划线]")
				}
				if run.FontSize > 0 {
					fmt.Printf(" [字号:%d]", run.FontSize)
				}
				fmt.Println()
			}
		}
	}
	fmt.Println()

	// 示例3: 获取文档中的所有表格
	fmt.Println("📄 示例3: 提取表格信息")
	tables, err := doc.GetTables()
	if err != nil {
		log.Printf("❌ 无法获取表格: %v", err)
		return
	}

	fmt.Printf("📊 文档包含 %d 个表格\n", len(tables))
	for i, table := range tables {
		fmt.Printf("   表格 %d: %d行 x %d列\n", i+1, len(table.Rows), table.Columns)
		
		// 显示表格内容（前几行）
		for rowIdx, row := range table.Rows {
			if rowIdx >= 3 { // 只显示前3行
				fmt.Printf("     ... (还有 %d 行)\n", len(table.Rows)-3)
				break
			}
			fmt.Printf("     行 %d: ", rowIdx+1)
			for colIdx, cell := range row.Cells {
				if colIdx > 0 {
					fmt.Print(" | ")
				}
				if cell.Text == "" {
					fmt.Print("(空)")
				} else {
					// 只显示前20个字符
					if len(cell.Text) > 20 {
						fmt.Printf("%s...", cell.Text[:20])
					} else {
						fmt.Print(cell.Text)
					}
				}
			}
			fmt.Println()
		}
	}
	fmt.Println()

	// 示例4: 获取文档结构信息
	fmt.Println("📄 示例4: 文档结构信息")
	container := doc.GetContainer()
	if container != nil {
		parts := container.GetParts()
		fmt.Printf("📁 文档包含 %d 个部分\n", len(parts))
		for uri, part := range parts {
			fmt.Printf("   %s (%d 字节)\n", uri, len(part.Data))
		}
	}
	fmt.Println()

	fmt.Println("✅ 基本使用示例完成")
	fmt.Println()
	fmt.Println("💡 更多示例请查看:")
	fmt.Println("   - examples/advanced_usage.go (高级用法)")
	fmt.Println("   - examples/document_modification.go (文档修改)")
	fmt.Println("   - examples/advanced_formatting.go (高级格式化)")
} 