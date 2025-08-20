package main

import (
	"fmt"

	"github.com/tanqiangyes/go-word/pkg/word"
)

func DemoDocumentParts() {
	fmt.Println("=== Go Word 文档部分支持演示 ===")

	// 演示文档部分功能
	demoDocumentParts()

	fmt.Println("文档部分演示完成！")
}

func demoDocumentParts() {
	fmt.Println("1. 创建文档部分结构...")

	// 创建新的文档部分
	parts := word.NewDocumentParts()

	// 添加页眉部分
	headerPart := word.HeaderPart{
		ID:   "header1",
		Type: word.FirstHeaderType,
		Content: []word.Paragraph{
			{
				Text:  "这是页眉内容",
				Style: "Header",
				Runs: []word.Run{
					{
						Text:     "这是页眉内容",
						FontSize: 12,
						FontName: "Arial",
					},
				},
			},
		},
		Properties: word.HeaderFooterProperties{
			DifferentFirst:   true,
			DifferentOddEven: false,
		},
	}
	parts.AddHeaderPart(headerPart)

	// 添加页脚部分
	footerPart := word.FooterPart{
		ID:   "footer1",
		Type: word.FirstFooterType,
		Content: []word.Paragraph{
			{
				Text:  "第 {PAGE} 页",
				Style: "Footer",
				Runs: []word.Run{
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
		Properties: word.HeaderFooterProperties{
			DifferentFirst:   true,
			DifferentOddEven: false,
		},
	}
	parts.AddFooterPart(footerPart)

	// 添加注释部分
	commentPart := word.CommentPart{
		ID: "comments1",
		Content: []word.Comment{
			{
				ID:       "comment1",
				Author:   "张三",
				Date:     "2024-01-01",
				Text:     "这是一个重要的注释",
				Initials: "ZS",
				Index:    1,
				Formatting: word.CommentFormatting{
					FontName:  "Arial",
					FontSize:  10,
					Bold:      false,
					Italic:    false,
					Color:     "000000",
					Highlight: "FFFF00",
				},
			},
		},
		Properties: word.CommentProperties{
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
	footnotePart := word.FootnotePart{
		ID: "footnotes1",
		Content: []word.Footnote{
			{
				ID:        "footnote1",
				Text:      "这是脚注内容",
				Number:    1,
				Type:      word.NormalFootnote,
				Reference: "ref1",
				Formatting: word.FootnoteFormatting{
					FontName: "Arial",
					FontSize: 9,
					Bold:     false,
					Italic:   false,
					Color:    "000000",
				},
			},
		},
		Properties: word.FootnoteProperties{
			RestartNumber: false,
			StartNumber:   1,
			NumberFormat:  "decimal",
			Position:      word.PageBottom,
			Layout: word.FootnoteLayout{
				Columns: 1,
				Spacing: 0.0,
				Indent:  0.0,
			},
		},
	}
	parts.AddFootnotePart(footnotePart)

	// 添加尾注部分
	endnotePart := word.EndnotePart{
		ID: "endnotes1",
		Content: []word.Endnote{
			{
				ID:        "endnote1",
				Text:      "这是尾注内容",
				Number:    1,
				Type:      word.NormalEndnote,
				Reference: "ref2",
				Formatting: word.EndnoteFormatting{
					FontName: "Arial",
					FontSize: 9,
					Bold:     false,
					Italic:   false,
					Color:    "000000",
				},
			},
		},
		Properties: word.EndnoteProperties{
			RestartNumber: false,
			StartNumber:   1,
			NumberFormat:  "decimal",
			Position:      word.DocumentEnd,
			Layout: word.EndnoteLayout{
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
	stylesPart := &word.StylesPart{
		ID: "styles1",
		Styles: word.StyleDefinitions{
			ParagraphStyles: []word.ParagraphStyle{
				{
					ID:      "Normal",
					Name:    "Normal",
					BasedOn: "",
					Next:    "Normal",
					Hidden:  false,
					Locked:  false,
				},
				{
					ID:      "Heading1",
					Name:    "Heading 1",
					BasedOn: "Normal",
					Next:    "Normal",
					Hidden:  false,
					Locked:  false,
				},
			},
			CharacterStyles: []word.CharacterStyle{
				{
					ID:             "DefaultParagraphFont",
					Name:           "Default Paragraph Font",
					BasedOn:        "",
					Next:           "",
					Link:           "",
					SemiHidden:     true,
					UnhideWhenUsed: true,
					QFormat:        true,
					Locked:         false,
					Properties: word.CharacterStyleProperties{
						Hidden:     false,
						Vanish:     false,
						SpecVanish: false,
					},
				},
			},
			DefaultStyles: word.DefaultStyleSet{
				Paragraph: "Normal",
				Character: "DefaultParagraphFont",
				Table:     "TableNormal",
				Numbering: "NoList",
				Properties: word.DefaultStyleProperties{
					Language: "zh-CN",
					Theme:    "Office",
					Hidden:   false,
				},
			},
			Properties: word.StyleProperties{
				Language: "zh-CN",
				Theme:    "Office",
				Hidden:   false,
			},
		},
	}
	parts.StylesPart = stylesPart

	fmt.Println("5. 演示设置部分...")

	// 创建设置部分
	settingsPart := &word.SettingsPart{
		ID: "settings1",
		Settings: word.DocumentSettings{
			ViewSettings: word.ViewSettings{
				ShowWhiteSpace:     true,
				ShowParagraphMarks: false,
				ShowHiddenText:     false,
				ShowBookmarks:      false,
				Zoom:               100,
				ZoomType:           "percent",
			},
			EditSettings: word.EditSettings{
				TrackChanges: false,
				Protection:   false,
				ReadOnly:     false,
				AutoFormat:   true,
				AutoCorrect:  true,
			},
			PrintSettings: word.PrintSettings{
				PrintBackground: true,
				PrintHiddenText: false,
				PrintComments:   true,
				PrintProperties: false,
			},
			OtherSettings: word.OtherSettings{
				UpdateFields:       true,
				EmbedTrueTypeFonts: true,
				SaveSubsetFonts:    true,
				Hidden:             false,
			},
		},
	}
	parts.SettingsPart = settingsPart

	fmt.Println("6. 最终文档部分摘要...")
	fmt.Println(parts.GetPartsSummary())

	fmt.Println("文档部分演示完成！")
}

func main() {
	DemoDocumentParts()
}
