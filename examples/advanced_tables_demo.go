package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func main() {
	fmt.Println("=== Go Word 高级表格功能演示 ===\n")

	// 演示高级表格功能
	demoAdvancedTables()

	fmt.Println("高级表格功能演示完成！")
}

func demoAdvancedTables() {
	fmt.Println("1. 创建高级表格系统...")
	
	// 创建高级表格系统
	tableSystem := wordprocessingml.NewAdvancedTableSystem()
	
	fmt.Println("2. 创建高级表格...")
	
	// 创建一个3x4的表格
	table := tableSystem.CreateAdvancedTable("示例表格", 3, 4)
	
	fmt.Printf("✓ 创建表格成功: %s\n", table.Name)
	fmt.Printf("表格ID: %s\n", table.ID)
	fmt.Printf("表格描述: %s\n", table.Description)
	fmt.Printf("行数: %d, 列数: %d\n", len(table.Rows), len(table.Columns))
	
	// 显示表格数据
	fmt.Println("3. 显示表格数据...")
	fmt.Println("表格数据:")
	for i, row := range table.Data {
		fmt.Printf("行 %d: %v\n", i+1, row)
	}
	
	// 显示表格结构
	fmt.Println("4. 显示表格结构...")
	fmt.Printf("表格属性:\n")
	fmt.Printf("  宽度: %.1f\n", table.Properties.Width)
	fmt.Printf("  高度: %.1f\n", table.Properties.Height)
	fmt.Printf("  对齐: %s\n", table.Properties.Alignment)
	fmt.Printf("  位置: %s\n", table.Properties.Position)
	
	fmt.Printf("表格布局:\n")
	fmt.Printf("  类型: %v\n", table.Properties.Layout.Type)
	fmt.Printf("  宽度: %.1f\n", table.Properties.Layout.Width)
	fmt.Printf("  对齐: %s\n", table.Properties.Layout.Alignment)
	fmt.Printf("  自动布局: %v\n", table.Properties.Layout.AutoLayout)
	
	// 显示行信息
	fmt.Println("5. 显示行信息...")
	for i, row := range table.Rows {
		fmt.Printf("行 %d:\n", i+1)
		fmt.Printf("  ID: %s\n", row.ID)
		fmt.Printf("  高度: %.1f\n", row.Height)
		fmt.Printf("  最小高度: %.1f\n", row.MinHeight)
		fmt.Printf("  最大高度: %.1f\n", row.MaxHeight)
		fmt.Printf("  隐藏: %v\n", row.Hidden)
		fmt.Printf("  重复表头: %v\n", row.RepeatHeader)
		fmt.Printf("  单元格数量: %d\n", len(row.Cells))
	}
	
	// 显示列信息
	fmt.Println("6. 显示列信息...")
	for j, col := range table.Columns {
		fmt.Printf("列 %d:\n", j+1)
		fmt.Printf("  ID: %s\n", col.ID)
		fmt.Printf("  宽度: %.1f\n", col.Width)
		fmt.Printf("  最小宽度: %.1f\n", col.MinWidth)
		fmt.Printf("  最大宽度: %.1f\n", col.MaxWidth)
		fmt.Printf("  隐藏: %v\n", col.Hidden)
		fmt.Printf("  自动适应: %v\n", col.AutoFit)
		fmt.Printf("  最佳适应: %v\n", col.BestFit)
	}
	
	// 显示表头信息
	fmt.Println("7. 显示表头信息...")
	for j, header := range table.Headers {
		fmt.Printf("表头 %d:\n", j+1)
		fmt.Printf("  ID: %s\n", header.ID)
		fmt.Printf("  标题: %s\n", header.Title)
		fmt.Printf("  级别: %d\n", header.Level)
		fmt.Printf("  样式: %s\n", header.Style)
		fmt.Printf("  对齐: %s\n", header.Alignment)
		fmt.Printf("  可见: %v\n", header.Visible)
		fmt.Printf("  重复: %v\n", header.Repeat)
	}
	
	// 显示单元格信息
	fmt.Println("8. 显示单元格信息...")
	for i, row := range table.Rows {
		for j, cell := range row.Cells {
			fmt.Printf("单元格 (%d,%d):\n", i+1, j+1)
			fmt.Printf("  ID: %s\n", cell.ID)
			fmt.Printf("  宽度: %.1f, 高度: %.1f\n", cell.Width, cell.Height)
			fmt.Printf("  对齐: %s, 垂直对齐: %s\n", cell.Alignment, cell.VerticalAlignment)
			fmt.Printf("  行跨度: %d, 列跨度: %d\n", cell.RowSpan, cell.ColSpan)
			fmt.Printf("  内容: %s\n", cell.Content.Text)
			fmt.Printf("  隐藏: %v, 锁定: %v\n", cell.Hidden, cell.Locked)
		}
	}
	
	fmt.Println("9. 演示单元格合并...")
	
	// 合并第一行的前两个单元格
	if err := tableSystem.MergeCells(table.ID, 0, 0, 0, 1); err != nil {
		fmt.Printf("合并单元格失败: %v\n", err)
	} else {
		fmt.Println("✓ 合并单元格成功")
	}
	
	// 显示合并后的表格数据
	fmt.Println("合并后的表格数据:")
	for i, row := range table.Data {
		fmt.Printf("行 %d: %v\n", i+1, row)
	}
	
	// 显示合并后的单元格信息
	fmt.Println("合并后的单元格信息:")
	for i, row := range table.Rows {
		for j, cell := range row.Cells {
			if !cell.Hidden {
				fmt.Printf("单元格 (%d,%d):\n", i+1, j+1)
				fmt.Printf("  行跨度: %d, 列跨度: %d\n", cell.RowSpan, cell.ColSpan)
				fmt.Printf("  内容: %s\n", cell.Content.Text)
				fmt.Printf("  隐藏: %v\n", cell.Hidden)
			}
		}
	}
	
	fmt.Println("10. 演示单元格拆分...")
	
	// 拆分合并的单元格
	if err := tableSystem.SplitCells(table.ID, 0, 0); err != nil {
		fmt.Printf("拆分单元格失败: %v\n", err)
	} else {
		fmt.Println("✓ 拆分单元格成功")
	}
	
	// 显示拆分后的表格数据
	fmt.Println("拆分后的表格数据:")
	for i, row := range table.Data {
		fmt.Printf("行 %d: %v\n", i+1, row)
	}
	
	// 显示拆分后的单元格信息
	fmt.Println("拆分后的单元格信息:")
	for i, row := range table.Rows {
		for j, cell := range row.Cells {
			fmt.Printf("单元格 (%d,%d):\n", i+1, j+1)
			fmt.Printf("  行跨度: %d, 列跨度: %d\n", cell.RowSpan, cell.ColSpan)
			fmt.Printf("  内容: %s\n", cell.Content.Text)
			fmt.Printf("  隐藏: %v\n", cell.Hidden)
		}
	}
	
	fmt.Println("11. 创建第二个表格...")
	
	// 创建第二个表格
	table2 := tableSystem.CreateAdvancedTable("数据表格", 5, 3)
	
	fmt.Printf("✓ 创建第二个表格成功: %s\n", table2.Name)
	fmt.Printf("表格ID: %s\n", table2.ID)
	fmt.Printf("行数: %d, 列数: %d\n", len(table2.Rows), len(table2.Columns))
	
	// 设置表格数据
	fmt.Println("12. 设置表格数据...")
	
	// 设置表头
	table2.Headers[0].Title = "姓名"
	table2.Headers[1].Title = "年龄"
	table2.Headers[2].Title = "部门"
	
	// 设置数据
	data := [][]string{
		{"张三", "25", "技术部"},
		{"李四", "30", "市场部"},
		{"王五", "28", "人事部"},
		{"赵六", "35", "财务部"},
		{"钱七", "27", "技术部"},
	}
	
	for i, row := range data {
		for j, value := range row {
			table2.Data[i][j] = value
			table2.Rows[i].Cells[j].Content.Text = value
		}
	}
	
	fmt.Println("设置后的表格数据:")
	for i, row := range table2.Data {
		fmt.Printf("行 %d: %v\n", i+1, row)
	}
	
	fmt.Println("13. 显示表格统计...")
	fmt.Println(tableSystem.GetTableSummary())
	
	fmt.Println("高级表格功能演示完成！")
} 