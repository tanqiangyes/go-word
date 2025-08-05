package examples

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/types"
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func DemoAdvancedStyles() {
	fmt.Println("=== Go Word 高级样式系统演示 ===\n")

	// 演示高级样式系统功能
	demoAdvancedStyleSystem()

	fmt.Println("高级样式系统演示完成！")
}

func demoAdvancedStyleSystem() {
	fmt.Println("1. 创建高级样式系统...")
	
	// 创建高级样式系统
	styleSystem := wordprocessingml.NewAdvancedStyleSystem()
	
	fmt.Println("2. 添加段落样式...")
	
	// 创建段落样式
	normalStyle := &wordprocessingml.ParagraphStyleDefinition{
		ID:          "Normal",
		Name:        "Normal",
		BasedOn:     "",
		Next:        "Normal",
		Link:        "",
		SemiHidden:     false,
		UnhideWhenUsed: true,
		QFormat:        true,
		Locked:         false,
		Properties: &wordprocessingml.ParagraphStyleProperties{
			Alignment: "left",
			KeepLines: true,
			KeepNext:  false,
			PageBreakBefore: false,
			WidowControl: true,
			OutlineLevel: 0,
			Font: &wordprocessingml.Font{
				Name: "Arial",
				Size: 11,
				Bold: false,
				Italic: false,
			},
		},
	}
	
	heading1Style := &wordprocessingml.ParagraphStyleDefinition{
		ID:          "Heading1",
		Name:        "Heading 1",
		BasedOn:     "Normal",
		Next:        "Normal",
		Link:        "",
		SemiHidden:     false,
		UnhideWhenUsed: true,
		QFormat:        true,
		Locked:         false,
		Properties: &wordprocessingml.ParagraphStyleProperties{
			Alignment: "left",
			KeepLines: true,
			KeepNext:  true,
			PageBreakBefore: false,
			WidowControl: true,
			OutlineLevel: 1,
			Font: &wordprocessingml.Font{
				Name: "Arial",
				Size: 16,
				Bold: true,
				Italic: false,
			},
		},
	}
	
	// 添加段落样式
	if err := styleSystem.AddParagraphStyle(normalStyle); err != nil {
		fmt.Printf("添加Normal样式失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加Normal样式成功")
	}
	
	if err := styleSystem.AddParagraphStyle(heading1Style); err != nil {
		fmt.Printf("添加Heading1样式失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加Heading1样式成功")
	}
	
	fmt.Println("3. 添加字符样式...")
	
	// 创建字符样式
	emphasisStyle := &wordprocessingml.CharacterStyleDefinition{
		ID:          "Emphasis",
		Name:        "Emphasis",
		BasedOn:     "DefaultParagraphFont",
		Link:        "",
		SemiHidden:     false,
		UnhideWhenUsed: true,
		QFormat:        true,
		Locked:         false,
		Properties: &wordprocessingml.CharacterStyleProperties{
			Font: &wordprocessingml.Font{
				Name: "Arial",
				Size: 11,
				Bold: false,
				Italic: true,
			},
			Hidden: false,
			Vanish: false,
			SpecVanish: false,
		},
	}
	
	strongStyle := &wordprocessingml.CharacterStyleDefinition{
		ID:          "Strong",
		Name:        "Strong",
		BasedOn:     "DefaultParagraphFont",
		Link:        "",
		SemiHidden:     false,
		UnhideWhenUsed: true,
		QFormat:        true,
		Locked:         false,
		Properties: &wordprocessingml.CharacterStyleProperties{
			Font: &wordprocessingml.Font{
				Name: "Arial",
				Size: 11,
				Bold: true,
				Italic: false,
			},
			Hidden: false,
			Vanish: false,
			SpecVanish: false,
		},
	}
	
	// 添加字符样式
	if err := styleSystem.AddCharacterStyle(emphasisStyle); err != nil {
		fmt.Printf("添加Emphasis样式失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加Emphasis样式成功")
	}
	
	if err := styleSystem.AddCharacterStyle(strongStyle); err != nil {
		fmt.Printf("添加Strong样式失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加Strong样式成功")
	}
	
	fmt.Println("4. 添加表格样式...")
	
	// 创建表格样式
	tableStyle := &wordprocessingml.TableStyleDefinition{
		ID:          "TableGrid",
		Name:        "Table Grid",
		BasedOn:     "",
		Next:        "",
		SemiHidden:     false,
		UnhideWhenUsed: true,
		QFormat:        true,
		Locked:         false,
		Properties: &wordprocessingml.TableStyleProperties{
			Alignment: "left",
			Hidden: false,
			AllowOverlap: false,
			AllowBreak: true,
			Borders: &wordprocessingml.TableBorders{
				Top:     wordprocessingml.BorderSide{Style: "single", Size: 1, Color: "000000"},
				Bottom:  wordprocessingml.BorderSide{Style: "single", Size: 1, Color: "000000"},
				Left:    wordprocessingml.BorderSide{Style: "single", Size: 1, Color: "000000"},
				Right:   wordprocessingml.BorderSide{Style: "single", Size: 1, Color: "000000"},
				InsideH: wordprocessingml.BorderSide{Style: "single", Size: 1, Color: "000000"},
				InsideV: wordprocessingml.BorderSide{Style: "single", Size: 1, Color: "000000"},
			},
		},
	}
	
	// 添加表格样式
	if err := styleSystem.AddTableStyle(tableStyle); err != nil {
		fmt.Printf("添加TableGrid样式失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加TableGrid样式成功")
	}
	
	fmt.Println("5. 演示样式查询...")
	
	// 查询样式
	if style := styleSystem.GetStyle("Normal"); style != nil {
		fmt.Printf("找到样式: %s, 类型: %v\n", style.Name, style.Type)
	}
	
	if style := styleSystem.GetStyle("Heading1"); style != nil {
		fmt.Printf("找到样式: %s, 类型: %v\n", style.Name, style.Type)
	}
	
	if style := styleSystem.GetStyle("Emphasis"); style != nil {
		fmt.Printf("找到样式: %s, 类型: %v\n", style.Name, style.Type)
	}
	
	if style := styleSystem.GetStyle("Strong"); style != nil {
		fmt.Printf("找到样式: %s, 类型: %v\n", style.Name, style.Type)
	}
	
	if style := styleSystem.GetStyle("TableGrid"); style != nil {
		fmt.Printf("找到样式: %s, 类型: %v\n", style.Name, style.Type)
	}
	
	fmt.Println("6. 演示继承链...")
	
	// 查看继承链
	if chain := styleSystem.GetInheritanceChain("Heading1"); len(chain) > 0 {
		fmt.Printf("Heading1继承链: %v\n", chain)
	}
	
	if chain := styleSystem.GetInheritanceChain("Emphasis"); len(chain) > 0 {
		fmt.Printf("Emphasis继承链: %v\n", chain)
	}
	
	fmt.Println("7. 演示样式应用...")
	
	// 创建测试内容
	paragraph := &types.Paragraph{
		Text:  "这是一个测试段落",
		Style: "",
		Runs: []types.Run{
			{
				Text:     "这是一个测试段落",
				FontSize: 11,
				FontName: "Arial",
			},
		},
	}
	
	run := &types.Run{
		Text:     "这是强调文本",
		FontSize: 11,
		FontName: "Arial",
	}
	
	table := &types.Table{
		Rows: []types.TableRow{
			{
				Cells: []types.TableCell{
					{Text: "单元格1"},
					{Text: "单元格2"},
				},
			},
		},
	}
	
	// 应用样式
	if err := styleSystem.ApplyStyle(paragraph, "Heading1"); err != nil {
		fmt.Printf("应用Heading1样式失败: %v\n", err)
	} else {
		fmt.Println("✓ 应用Heading1样式成功")
		fmt.Printf("段落样式: %s\n", paragraph.Style)
	}
	
	if err := styleSystem.ApplyStyle(run, "Emphasis"); err != nil {
		fmt.Printf("应用Emphasis样式失败: %v\n", err)
	} else {
		fmt.Println("✓ 应用Emphasis样式成功")
		fmt.Printf("运行字体: %s, 大小: %d, 斜体: %v\n", run.FontName, run.FontSize, run.Italic)
	}
	
	if err := styleSystem.ApplyStyle(table, "TableGrid"); err != nil {
		fmt.Printf("应用TableGrid样式失败: %v\n", err)
	} else {
		fmt.Println("✓ 应用TableGrid样式成功")
	}
	
	fmt.Println("8. 演示样式冲突解决...")
	
	// 尝试添加同名样式（冲突）
	conflictStyle := &wordprocessingml.ParagraphStyleDefinition{
		ID:          "Normal",
		Name:        "Normal",
		BasedOn:     "",
		Next:        "Normal",
		Link:        "",
		SemiHidden:     false,
		UnhideWhenUsed: true,
		QFormat:        true,
		Locked:         false,
		Properties: &wordprocessingml.ParagraphStyleProperties{
			Alignment: "center",
			KeepLines: true,
			KeepNext:  false,
			PageBreakBefore: false,
			WidowControl: true,
			OutlineLevel: 0,
		},
	}
	
	if err := styleSystem.AddParagraphStyle(conflictStyle); err != nil {
		fmt.Printf("样式冲突处理: %v\n", err)
	} else {
		fmt.Println("✓ 冲突样式添加成功")
	}
	
	fmt.Println("9. 显示样式系统摘要...")
	fmt.Println(styleSystem.GetStyleSummary())
	
	fmt.Println("高级样式系统演示完成！")
} 