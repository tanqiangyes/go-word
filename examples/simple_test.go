package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/types"
)

func main() {
	// 测试基本类型创建
	paragraph := types.Paragraph{
		Text:  "这是一个测试段落",
		Style: "Normal",
		Runs: []types.Run{
			{
				Text:     "这是测试文本",
				Bold:     true,
				Italic:   false,
				FontSize: 12,
				FontName: "Arial",
			},
		},
	}

	table := types.Table{
		Rows: []types.TableRow{
			{
				Cells: []types.TableCell{
					{Text: "单元格1"},
					{Text: "单元格2"},
				},
			},
		},
		Columns: 2,
	}

	content := types.DocumentContent{
		Paragraphs: []types.Paragraph{paragraph},
		Tables:     []types.Table{table},
		Text:       "测试文档内容",
	}

	fmt.Println("✅ 基本类型创建成功")
	fmt.Printf("段落数量: %d\n", len(content.Paragraphs))
	fmt.Printf("表格数量: %d\n", len(content.Tables))
	fmt.Printf("文档文本: %s\n", content.Text)

	log.Println("基本功能测试完成")
} 