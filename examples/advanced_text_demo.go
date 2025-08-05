package examples

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func DemoAdvancedText() {
	fmt.Println("=== Go Word 高级文本功能演示 ===\n")

	// 演示高级文本功能
	demoAdvancedText()

	fmt.Println("高级文本功能演示完成！")
}

func demoAdvancedText() {
	fmt.Println("1. 创建高级文本系统...")
	
	// 创建高级文本系统
	textSystem := wordprocessingml.NewAdvancedTextSystem()
	
	fmt.Println("2. 创建高级文本...")
	
	// 创建文本内容
	content := `这是一个示例文档。

它包含多个段落，每个段落都有不同的格式和样式。

这个段落包含一些特殊格式的文本，比如粗体、斜体等。

最后，这个文档展示了高级文本处理功能。`
	
	// 创建高级文本
	text := textSystem.CreateAdvancedText("示例文档", content)
	
	fmt.Printf("✓ 创建文本成功: %s\n", text.Name)
	fmt.Printf("文本ID: %s\n", text.ID)
	fmt.Printf("文本描述: %s\n", text.Description)
	fmt.Printf("内容长度: %d 字符\n", len(text.Content))
	
	// 显示文本内容
	fmt.Println("3. 显示文本内容...")
	fmt.Println("原始内容:")
	fmt.Println(text.Content)
	
	// 显示格式化内容
	fmt.Println("4. 显示格式化内容...")
	fmt.Printf("段落数量: %d\n", len(text.FormattedContent.Paragraphs))
	fmt.Printf("语言: %s\n", text.FormattedContent.Language)
	fmt.Printf("方向: %v\n", text.FormattedContent.Direction)
	
	// 显示段落信息
	fmt.Println("5. 显示段落信息...")
	for i, paragraph := range text.FormattedContent.Paragraphs {
		fmt.Printf("段落 %d:\n", i+1)
		fmt.Printf("  ID: %s\n", paragraph.ID)
		fmt.Printf("  索引: %d\n", paragraph.Index)
		fmt.Printf("  文本: %s\n", paragraph.Text)
		fmt.Printf("  运行数量: %d\n", len(paragraph.Runs))
		fmt.Printf("  隐藏: %v\n", paragraph.Hidden)
		fmt.Printf("  锁定: %v\n", paragraph.Locked)
		
		// 显示段落属性
		if paragraph.Properties != nil {
			fmt.Printf("  对齐: %s\n", paragraph.Properties.Alignment)
			fmt.Printf("  两端对齐: %s\n", paragraph.Properties.Justification)
			fmt.Printf("  保持行: %v\n", paragraph.Properties.KeepLines)
			fmt.Printf("  保持下一段: %v\n", paragraph.Properties.KeepNext)
			fmt.Printf("  段前分页: %v\n", paragraph.Properties.PageBreakBefore)
			fmt.Printf("  孤行控制: %v\n", paragraph.Properties.WidowControl)
		}
		
		// 显示段落样式
		if paragraph.Style != nil {
			fmt.Printf("  样式ID: %s\n", paragraph.Style.ID)
			fmt.Printf("  样式名称: %s\n", paragraph.Style.Name)
			fmt.Printf("  基于: %s\n", paragraph.Style.BasedOn)
			fmt.Printf("  下一个: %s\n", paragraph.Style.Next)
		}
		
		// 显示运行信息
		for j, run := range paragraph.Runs {
			fmt.Printf("    运行 %d:\n", j+1)
			fmt.Printf("      ID: %s\n", run.ID)
			fmt.Printf("      索引: %d\n", run.Index)
			fmt.Printf("      文本: %s\n", run.Text)
			fmt.Printf("      隐藏: %v\n", run.Hidden)
			fmt.Printf("      锁定: %v\n", run.Locked)
			
			// 显示运行属性
			if run.Properties != nil {
				if run.Properties.Font != nil {
					fmt.Printf("      字体名称: %s\n", run.Properties.Font.Name)
					fmt.Printf("      字体大小: %.1f\n", run.Properties.Font.Size)
					fmt.Printf("      字体族: %s\n", run.Properties.Font.Family)
					fmt.Printf("      粗体: %v\n", run.Properties.Font.Bold)
					fmt.Printf("      斜体: %v\n", run.Properties.Font.Italic)
					fmt.Printf("      下划线: %v\n", run.Properties.Font.Underline)
					fmt.Printf("      删除线: %v\n", run.Properties.Font.Strike)
					fmt.Printf("      颜色: %s\n", run.Properties.Font.Color)
					fmt.Printf("      高亮: %s\n", run.Properties.Font.Highlight)
					fmt.Printf("      上标: %v\n", run.Properties.Font.Superscript)
					fmt.Printf("      下标: %v\n", run.Properties.Font.Subscript)
					fmt.Printf("      小型大写: %v\n", run.Properties.Font.SmallCaps)
					fmt.Printf("      全大写: %v\n", run.Properties.Font.AllCaps)
				}
				
				if run.Properties.Format != nil {
					fmt.Printf("      对齐: %s\n", run.Properties.Format.Alignment)
					fmt.Printf("      两端对齐: %s\n", run.Properties.Format.Justification)
					fmt.Printf("      保持行: %v\n", run.Properties.Format.KeepLines)
					fmt.Printf("      保持下一段: %v\n", run.Properties.Format.KeepNext)
				}
				
				if run.Properties.Position != nil {
					fmt.Printf("      位置X: %.1f\n", run.Properties.Position.X)
					fmt.Printf("      位置Y: %.1f\n", run.Properties.Position.Y)
					fmt.Printf("      位置Z: %.1f\n", run.Properties.Position.Z)
					fmt.Printf("      相对位置: %v\n", run.Properties.Position.Relative)
					fmt.Printf("      绝对位置: %v\n", run.Properties.Position.Absolute)
				}
				
				fmt.Printf("      语言: %s\n", run.Properties.Language)
				fmt.Printf("      方向: %v\n", run.Properties.Direction)
			}
			
			// 显示运行样式
			if run.Style != nil {
				fmt.Printf("      样式ID: %s\n", run.Style.ID)
				fmt.Printf("      样式名称: %s\n", run.Style.Name)
				fmt.Printf("      基于: %s\n", run.Style.BasedOn)
				fmt.Printf("      下一个: %s\n", run.Style.Next)
				fmt.Printf("      隐藏: %v\n", run.Style.Hidden)
				fmt.Printf("      锁定: %v\n", run.Style.Locked)
			}
			
			// 显示运行效果
			fmt.Printf("      效果数量: %d\n", len(run.Effects))
		}
	}
	
	// 显示文本属性
	fmt.Println("6. 显示文本属性...")
	if text.Properties != nil {
		fmt.Printf("语言: %s\n", text.Properties.Language)
		fmt.Printf("方向: %v\n", text.Properties.Direction)
		fmt.Printf("阅读顺序: %v\n", text.Properties.ReadingOrder)
		fmt.Printf("受保护: %v\n", text.Properties.Protected)
		fmt.Printf("可编辑: %v\n", text.Properties.Editable)
		
		if text.Properties.Format != nil {
			fmt.Printf("格式类型: %s\n", text.Properties.Format.Type)
			fmt.Printf("格式版本: %s\n", text.Properties.Format.Version)
			fmt.Printf("兼容性: %v\n", text.Properties.Format.Compatible)
			fmt.Printf("可扩展: %v\n", text.Properties.Format.Extensible)
		}
		
		if text.Properties.Layout != nil {
			fmt.Printf("布局类型: %s\n", text.Properties.Layout.Type)
			fmt.Printf("布局流: %s\n", text.Properties.Layout.Flow)
			fmt.Printf("自动换行: %v\n", text.Properties.Layout.Wrapping)
			fmt.Printf("溢出处理: %s\n", text.Properties.Layout.Overflow)
		}
	}
	
	// 显示文本效果
	fmt.Println("7. 显示文本效果...")
	fmt.Printf("效果数量: %d\n", len(text.Effects))
	for i, effect := range text.Effects {
		fmt.Printf("效果 %d:\n", i+1)
		fmt.Printf("  ID: %s\n", effect.ID)
		fmt.Printf("  名称: %s\n", effect.Name)
		fmt.Printf("  类型: %v\n", effect.Type)
		fmt.Printf("  启用: %v\n", effect.Enabled)
		fmt.Printf("  持续时间: %.1f\n", effect.Duration)
		
		if effect.Properties != nil {
			fmt.Printf("  颜色: %s\n", effect.Properties.Color)
			fmt.Printf("  大小: %.1f\n", effect.Properties.Size)
			fmt.Printf("  透明度: %.1f\n", effect.Properties.Opacity)
			fmt.Printf("  位置X: %.1f\n", effect.Properties.X)
			fmt.Printf("  位置Y: %.1f\n", effect.Properties.Y)
			fmt.Printf("  位置Z: %.1f\n", effect.Properties.Z)
			fmt.Printf("  模糊: %.1f\n", effect.Properties.Blur)
			fmt.Printf("  距离: %.1f\n", effect.Properties.Distance)
			fmt.Printf("  角度: %.1f\n", effect.Properties.Angle)
		}
	}
	
	fmt.Println("8. 演示文本效果应用...")
	
	// 应用发光效果
	glowProperties := &wordprocessingml.EffectProperties{
		Color:   "FFFF00",
		Size:    5.0,
		Opacity: 0.8,
		X:       0.0,
		Y:       0.0,
		Z:       0.0,
		Blur:    2.0,
		Distance: 0.0,
		Angle:   0.0,
	}
	
	if err := textSystem.ApplyTextEffect(text.ID, wordprocessingml.GlowEffect, glowProperties); err != nil {
		fmt.Printf("应用发光效果失败: %v\n", err)
	} else {
		fmt.Println("✓ 应用发光效果成功")
	}
	
	// 应用阴影效果
	shadowProperties := &wordprocessingml.EffectProperties{
		Color:   "808080",
		Size:    3.0,
		Opacity: 0.6,
		X:       2.0,
		Y:       2.0,
		Z:       0.0,
		Blur:    1.0,
		Distance: 2.0,
		Angle:   45.0,
	}
	
	if err := textSystem.ApplyTextEffect(text.ID, wordprocessingml.ShadowEffect, shadowProperties); err != nil {
		fmt.Printf("应用阴影效果失败: %v\n", err)
	} else {
		fmt.Println("✓ 应用阴影效果成功")
	}
	
	// 显示应用效果后的信息
	fmt.Printf("应用效果后的效果数量: %d\n", len(text.Effects))
	
	fmt.Println("9. 创建第二个文本...")
	
	// 创建第二个文本
	content2 := `这是第二个示例文档。

它展示了不同的文本格式和样式。

包含一些特殊字符和格式。`
	
	text2 := textSystem.CreateAdvancedText("第二个文档", content2)
	
	fmt.Printf("✓ 创建第二个文本成功: %s\n", text2.Name)
	fmt.Printf("文本ID: %s\n", text2.ID)
	fmt.Printf("内容长度: %d 字符\n", len(text2.Content))
	
	// 显示文本统计
	fmt.Println("10. 显示文本统计...")
	fmt.Println(textSystem.GetTextSummary())
	
	fmt.Println("高级文本功能演示完成！")
} 