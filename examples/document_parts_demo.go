package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func main() {
	fmt.Println("=== Go Word 文档部分支持演示 ===\n")

	// 演示文档部分功能
	demoDocumentParts()

	fmt.Println("文档部分演示完成！")
}

func demoDocumentParts() {
	fmt.Println("1. 创建文档部分结构...")
	
	// 创建新的文档部分
	parts := wordprocessingml.NewDocumentParts()
	
	// 添加页眉部分
	headerPart := wordprocessingml.HeaderPart{
		ID:   "header1",
		Type: wordprocessingml.FirstHeader,
		Content: []wordprocessingml.Paragraph{
			{
				Text:  "这是页眉内容",
				Style: "Header",
				Runs: []wordprocessingml.Run{
					{
						Text:     "这是页眉内容",
						FontSize: 12,
						FontName: "Arial",
					},
				},
			},
		},
		Properties: wordprocessingml.HeaderFooterProperties{
			DifferentFirst: true,
			DifferentOddEven: false,
		},
	}
	parts.AddHeaderPart(headerPart)
	
	// 添加页脚部分
	footerPart := wordprocessingml.FooterPart{
		ID:   "footer1",
		Type: wordprocessingml.FirstFooter,
		Content: []wordprocessingml.Paragraph{
			{
				Text:  "第 {PAGE} 页",
				Style: "Footer",
				Runs: []wordprocessingml.Run{
					{
						Text:     "第 ",
						FontSize: 10,
						FontName: "Arial",
					},
					{
						Text:     "{PAGE}",
						FontSize: 10,
						FontName: "Arial",
						Bold:     true,
					},
					{
						Text:     " 页",
						FontSize: 10,
						FontName: "Arial",
					},
				},
			},
		},
		Properties: wordprocessingml.HeaderFooterProperties{
			DifferentFirst: true,
			DifferentOddEven: false,
		},
	}
	parts.AddFooterPart(footerPart)
	
	// 添加注释部分
	commentPart := wordprocessingml.CommentPart{
		ID: "comments1",
		Content: []wordprocessingml.Comment{
			{
				ID:       "comment1",
				Author:   "张三",
				Date:     "2024-01-01",
				Text:     "这是一个重要的注释",
				Initials: "ZS",
				Index:    1,
				Formatting: wordprocessingml.CommentFormatting{
					FontName:  "Arial",
					FontSize:  10,
					Bold:      false,
					Italic:    false,
					Color:     "000000",
					Highlight: "FFFF00",
				},
			},
		},
		Properties: wordprocessingml.CommentProperties{
			Visible:    true,
			Locked:     false,
			Resolved:   false,
			ShowAuthor: true,
			ShowDate:   true,
			ShowTime:   true,
		},
	}
	parts.AddCommentPart(commentPart)
	
	// 添加脚注部分
	footnotePart := wordprocessingml.FootnotePart{
		ID: "footnotes1",
		Content: []wordprocessingml.Footnote{
			{
				ID:       "footnote1",
				Text:     "这是脚注内容",
				Number:   1,
				Type:     wordprocessingml.NormalFootnote,
				Reference: "ref1",
				Formatting: wordprocessingml.FootnoteFormatting{
					FontName: "Arial",
					FontSize: 9,
					Bold:     false,
					Italic:   false,
					Color:    "000000",
				},
			},
		},
		Properties: wordprocessingml.FootnoteProperties{
			RestartNumber: false,
			StartNumber:   1,
			NumberFormat:  "decimal",
			Position:      wordprocessingml.PageBottom,
			Layout: wordprocessingml.FootnoteLayout{
				Columns: 1,
				Spacing: 0.0,
				Indent:  0.0,
			},
		},
	}
	parts.AddFootnotePart(footnotePart)
	
	// 添加尾注部分
	endnotePart := wordprocessingml.EndnotePart{
		ID: "endnotes1",
		Content: []wordprocessingml.Endnote{
			{
				ID:       "endnote1",
				Text:     "这是尾注内容",
				Number:   1,
				Type:     wordprocessingml.NormalEndnote,
				Reference: "ref2",
				Formatting: wordprocessingml.EndnoteFormatting{
					FontName: "Arial",
					FontSize: 9,
					Bold:     false,
					Italic:   false,
					Color:    "000000",
				},
			},
		},
		Properties: wordprocessingml.EndnoteProperties{
			RestartNumber: false,
			StartNumber:   1,
			NumberFormat:  "decimal",
			Position:      wordprocessingml.DocumentEnd,
			Layout: wordprocessingml.EndnoteLayout{
				Columns: 1,
				Spacing: 0.0,
				Indent:  0.0,
			},
		},
	}
	parts.AddEndnotePart(endnotePart)
	
	fmt.Println("2. 演示文档部分查询...")
	
	// 查询页眉部分
	if header := parts.GetHeaderPart("header1"); header != nil {
		fmt.Printf("找到页眉: %s, 类型: %v\n", header.ID, header.Type)
	}
	
	// 查询页脚部分
	if footer := parts.GetFooterPart("footer1"); footer != nil {
		fmt.Printf("找到页脚: %s, 类型: %v\n", footer.ID, footer.Type)
	}
	
	// 查询注释部分
	if comment := parts.GetCommentPart("comments1"); comment != nil {
		fmt.Printf("找到注释部分: %s, 注释数量: %d\n", comment.ID, len(comment.Content))
	}
	
	// 查询脚注部分
	if footnote := parts.GetFootnotePart("footnotes1"); footnote != nil {
		fmt.Printf("找到脚注部分: %s, 脚注数量: %d\n", footnote.ID, len(footnote.Content))
	}
	
	// 查询尾注部分
	if endnote := parts.GetEndnotePart("endnotes1"); endnote != nil {
		fmt.Printf("找到尾注部分: %s, 尾注数量: %d\n", endnote.ID, len(endnote.Content))
	}
	
	fmt.Println("3. 显示文档部分摘要...")
	fmt.Println(parts.GetPartsSummary())
	
	fmt.Println("4. 演示样式部分...")
	
	// 创建样式部分
	stylesPart := &wordprocessingml.StylesPart{
		ID: "styles1",
		Styles: wordprocessingml.StyleDefinitions{
			ParagraphStyles: []wordprocessingml.ParagraphStyle{
				{
					ID:          "Normal",
					Name:        "Normal",
					BasedOn:     "",
					Next:        "Normal",
					Link:        "",
					SemiHidden:     false,
					UnhideWhenUsed: true,
					QFormat:        true,
					Locked:         false,
					Properties: wordprocessingml.ParagraphStyleProperties{
						Alignment: "left",
						KeepLines: true,
						KeepNext:  false,
						PageBreakBefore: false,
						WidowControl: true,
					},
				},
				{
					ID:          "Heading1",
					Name:        "Heading 1",
					BasedOn:     "Normal",
					Next:        "Normal",
					Link:        "",
					SemiHidden:     false,
					UnhideWhenUsed: true,
					QFormat:        true,
					Locked:         false,
					Properties: wordprocessingml.ParagraphStyleProperties{
						Alignment: "left",
						KeepLines: true,
						KeepNext:  true,
						PageBreakBefore: false,
						WidowControl: true,
					},
				},
			},
			CharacterStyles: []wordprocessingml.CharacterStyle{
				{
					ID:          "DefaultParagraphFont",
					Name:        "Default Paragraph Font",
					BasedOn:     "",
					Link:        "",
					SemiHidden:     true,
					UnhideWhenUsed: true,
					QFormat:        true,
					Locked:         false,
					Properties: wordprocessingml.CharacterStyleProperties{
						Hidden: false,
						Vanish: false,
						SpecVanish: false,
					},
				},
			},
			DefaultStyles: wordprocessingml.DefaultStyleSet{
				Paragraph: "Normal",
				Character: "DefaultParagraphFont",
				Table:     "TableNormal",
				Numbering: "NoList",
				Properties: wordprocessingml.DefaultStyleProperties{
					Language: "zh-CN",
					Theme:    "Office",
					Hidden:   false,
				},
			},
			Properties: wordprocessingml.StyleProperties{
				Language: "zh-CN",
				Theme:    "Office",
				Hidden:   false,
			},
		},
	}
	parts.StylesPart = stylesPart
	
	fmt.Println("5. 演示设置部分...")
	
	// 创建设置部分
	settingsPart := &wordprocessingml.SettingsPart{
		ID: "settings1",
		Settings: wordprocessingml.DocumentSettings{
			ViewSettings: wordprocessingml.ViewSettings{
				ShowWhiteSpace:     true,
				ShowParagraphMarks: false,
				ShowHiddenText:     false,
				ShowBookmarks:      false,
				Zoom:               100,
				ZoomType:           "percent",
			},
			EditSettings: wordprocessingml.EditSettings{
				TrackChanges: false,
				Protection:   false,
				ReadOnly:     false,
				AutoFormat:   true,
				AutoCorrect:  true,
			},
			PrintSettings: wordprocessingml.PrintSettings{
				PrintBackground:  true,
				PrintHiddenText:  false,
				PrintComments:    true,
				PrintProperties:  false,
			},
			OtherSettings: wordprocessingml.OtherSettings{
				UpdateFields:        true,
				EmbedTrueTypeFonts:  true,
				SaveSubsetFonts:     true,
				Hidden:              false,
			},
		},
	}
	parts.SettingsPart = settingsPart
	
	fmt.Println("6. 最终文档部分摘要...")
	fmt.Println(parts.GetPartsSummary())
	
	fmt.Println("文档部分演示完成！")
} 