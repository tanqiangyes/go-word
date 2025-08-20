package tests

import (
    "strings"
    "testing"
    "time"

    "github.com/tanqiangyes/go-word/pkg/wordprocessingml"
    "github.com/tanqiangyes/go-word/pkg/types"
)

func TestNewDocumentQualityManager(t *testing.T) {
    // Create a document with proper structure
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "测试段落", Runs: []types.Run{{Text: "测试文本", FontName: "Arial", FontSize: 12}}},
            },
            Tables: []types.Table{},
        },
    })

    manager := word.NewDocumentQualityManager(doc)

    if manager == nil {
        t.Fatal("DocumentQualityManager should not be nil")
    }

    if manager.Document != doc {
        t.Error("Document should be set correctly")
    }

    if manager.Settings == nil {
        t.Fatal("Settings should not be nil")
    }

    if manager.Metrics == nil {
        t.Fatal("Metrics should not be nil")
    }

    // 检查默认设置
    if !manager.Settings.EnableMetadataManagement {
        t.Error("EnableMetadataManagement should be true by default")
    }

    if !manager.Settings.EnableContentQuality {
        t.Error("EnableContentQuality should be true by default")
    }
}

func TestImproveDocumentQuality(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "测试段落，包含一些格式问题。", Runs: []types.Run{{Text: "测试文本", FontName: "", FontSize: 0}}},
            },
            Tables: []types.Table{},
        },
    })

    manager := word.NewDocumentQualityManager(doc)

    err := manager.ImproveDocumentQuality()
    if err != nil {
        t.Fatalf("ImproveDocumentQuality failed: %v", err)
    }

    // 检查处理时间
    if manager.Metrics.ProcessingTime == 0 {
        t.Error("ProcessingTime should be set")
    }

    // 检查最后更新时间
    if manager.Metrics.LastUpdated.IsZero() {
        t.Error("LastUpdated should be set")
    }
}

func TestGenerateMetadata(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "第一个段落"},
                {Text: "第二个段落"},
            },
            Tables: []types.Table{
                {Rows: []types.TableRow{{Cells: []types.TableCell{{Text: "单元格"}}}}},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    metadata := manager.GenerateMetadata()

    if metadata == nil {
        t.Fatal("Metadata should not be nil")
    }

    if metadata.Created.IsZero() {
        t.Error("Created time should be set")
    }

    if metadata.Modified.IsZero() {
        t.Error("Modified time should be set")
    }

    if metadata.Creator != "Go-Word Library" {
        t.Error("Creator should be set correctly")
    }

    if metadata.Paragraphs != 2 {
        t.Errorf("Expected 2 paragraphs, got %d", metadata.Paragraphs)
    }

    if metadata.Tables != 1 {
        t.Errorf("Expected 1 table, got %d", metadata.Tables)
    }

    if metadata.Words != 2 {
        t.Errorf("Expected 2 words, got %d", metadata.Words)
    }
}

func TestValidateMetadata(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        DocumentProperties: make(map[string]interface{}),
    })

    manager := word.NewDocumentQualityManager(doc)

    err := manager.ValidateMetadata()
    if err != nil {
        t.Fatalf("ValidateMetadata failed: %v", err)
    }

    // 检查是否设置了默认值
    mainPart := doc.GetMainPart()
    if mainPart.DocumentProperties["title"] != "Untitled Document" {
        t.Error("Default title should be set")
    }

    if mainPart.DocumentProperties["author"] != "Unknown Author" {
        t.Error("Default author should be set")
    }

    if mainPart.DocumentProperties["language"] != "zh-CN" {
        t.Error("Default language should be set")
    }
}

func TestNormalizeText(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "  测试文本，包含多余空格  ", Runs: []types.Run{{Text: "  运行文本  "}}},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.NormalizeText(doc.GetMainPart().Content)

    // 检查空格是否被标准化
    if doc.GetMainPart().Content.Paragraphs[0].Text != "测试文本,包含多余空格" {
        t.Errorf("Expected normalized text, got: %s", doc.GetMainPart().Content.Paragraphs[0].Text)
    }

    if doc.GetMainPart().Content.Paragraphs[0].Runs[0].Text != "运行文本" {
        t.Errorf("Expected normalized run text, got: %s", doc.GetMainPart().Content.Paragraphs[0].Runs[0].Text)
    }
}

func TestNormalizePunctuation(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "测试文本，使用中文标点符号。！？"},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    normalized := manager.NormalizePunctuation(doc.GetMainPart().Content.Paragraphs[0].Text)

    expected := "测试文本,使用中文标点符号.!?"
    if normalized != expected {
        t.Errorf("Expected %s, got %s", expected, normalized)
    }
}

func TestNormalizeCase(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "test sentence"},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    normalized := manager.NormalizeCase(doc.GetMainPart().Content.Paragraphs[0].Text)

    expected := "Test sentence"
    if normalized != expected {
        t.Errorf("Expected %s, got %s", expected, normalized)
    }
}

func TestCheckSpelling(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "teh recieve seperate"},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.CheckSpelling(doc.GetMainPart().Content)

    expected := "the receive separate"
    if doc.GetMainPart().Content.Paragraphs[0].Text != expected {
        t.Errorf("Expected %s, got %s", expected, doc.GetMainPart().Content.Paragraphs[0].Text)
    }
}

func TestCheckGrammar(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "test sentence"},
                {Text: "another sentence"},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.CheckGrammar(doc.GetMainPart().Content)

    // 检查首字母大写
    if doc.GetMainPart().Content.Paragraphs[0].Text != "Test sentence." {
        t.Errorf("Expected 'Test sentence.', got %s", doc.GetMainPart().Content.Paragraphs[0].Text)
    }

    if doc.GetMainPart().Content.Paragraphs[1].Text != "Another sentence." {
        t.Errorf("Expected 'Another sentence.', got %s", doc.GetMainPart().Content.Paragraphs[1].Text)
    }
}

func TestCheckConsistency(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "段落1", Runs: []types.Run{{Text: "文本1", FontName: "Arial", FontSize: 12}}},
                {Text: "段落2", Runs: []types.Run{{Text: "文本2", FontName: "", FontSize: 0}}},
                {Text: "段落3", Runs: []types.Run{{Text: "文本3", FontName: "Times New Roman", FontSize: 14}}},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.CheckConsistency(doc.GetMainPart().Content)

    // 检查是否应用了最常见的字体
    if doc.GetMainPart().Content.Paragraphs[1].Runs[0].FontName != "Arial" {
        t.Errorf("Expected Arial font, got %s", doc.GetMainPart().Content.Paragraphs[1].Runs[0].FontName)
    }
}

func TestRemoveEmptyElements(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "有效段落"},
                {Text: "   "}, // 空段落
                {Text: "另一个有效段落"},
                {Text: "", Runs: []types.Run{}}, // 空段落
            },
            Tables: []types.Table{
                {Rows: []types.TableRow{{Cells: []types.TableCell{{Text: "有效表格"}}}}},
                {Rows: []types.TableRow{}}, // 空表格
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.RemoveEmptyElements(doc.GetMainPart().Content)

    // 检查空段落是否被移除
    if len(doc.GetMainPart().Content.Paragraphs) != 2 {
        t.Errorf("Expected 2 paragraphs, got %d", len(doc.GetMainPart().Content.Paragraphs))
    }

    // 检查空表格是否被移除
    if len(doc.GetMainPart().Content.Tables) != 1 {
        t.Errorf("Expected 1 table, got %d", len(doc.GetMainPart().Content.Tables))
    }
}

func TestAutoFixStructure(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Tables: []types.Table{
                {
                    Rows: []types.TableRow{
                        {Cells: []types.TableCell{{Text: "单元格1"}, {Text: "单元格2"}}},
                        {Cells: []types.TableCell{{Text: "单元格3"}}}, // 缺少一个单元格
                    },
                },
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.AutoFixStructure(doc.GetMainPart().Content)

    // 检查表格结构是否被修复
    table := doc.GetMainPart().Content.Tables[0]
    if table.Columns != 2 {
        t.Errorf("Expected 2 columns, got %d", table.Columns)
    }

    if len(table.Rows[1].Cells) != 2 {
        t.Errorf("Expected 2 cells in second row, got %d", len(table.Rows[1].Cells))
    }

    if table.Rows[1].Cells[1].Text != "" {
        t.Errorf("Expected empty cell, got %s", table.Rows[1].Cells[1].Text)
    }
}

func TestApplyConsistentFormatting(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "段落1", Runs: []types.Run{{Text: "文本1", FontName: "", FontSize: 0, Color: ""}}},
                {Text: "段落2", Runs: []types.Run{{Text: "文本2", FontName: "Arial", FontSize: 12, Color: "#000000"}}},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.ApplyConsistentFormatting(doc.GetMainPart().Content)

    // 检查默认格式是否被应用
    run := doc.GetMainPart().Content.Paragraphs[0].Runs[0]
    if run.FontName != "Arial" {
        t.Errorf("Expected Arial font, got %s", run.FontName)
    }

    if run.FontSize != 12 {
        t.Errorf("Expected font size 12, got %d", run.FontSize)
    }

    if run.Color != "#000000" {
        t.Errorf("Expected color #000000, got %s", run.Color)
    }
}

func TestCheckThemeConsistency(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "段落1", Runs: []types.Run{{Text: "文本1", Color: "#000000"}}},
                {Text: "段落2", Runs: []types.Run{{Text: "文本2", Color: "#000000"}}},
                {Text: "段落3", Runs: []types.Run{{Text: "文本3", Color: "#FF0000"}}},
                {Text: "段落4", Runs: []types.Run{{Text: "文本4", Color: ""}}},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.CheckThemeConsistency(doc.GetMainPart().Content)

    // 检查是否应用了最常见的颜色
    if doc.GetMainPart().Content.Paragraphs[3].Runs[0].Color != "#000000" {
        t.Errorf("Expected color #000000, got %s", doc.GetMainPart().Content.Paragraphs[3].Runs[0].Color)
    }
}

func TestGenerateAltText(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Tables: []types.Table{
                {
                    Rows: []types.TableRow{
                        {Cells: []types.TableCell{{Text: "标题1"}, {Text: "标题2"}}},
                        {Cells: []types.TableCell{{Text: "数据1"}, {Text: "数据2"}}},
                    },
                },
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.GenerateAltText(doc.GetMainPart().Content)

    // 这个测试主要是确保函数不会崩溃
    // 实际的替代文本生成逻辑可以在后续版本中实现
}

func TestAddStructureTags(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "标题段落", Style: ""},
                {Text: "普通段落", Style: ""},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)
    manager.AddStructureTags(doc.GetMainPart().Content)

    // 检查标题样式是否被应用
    if doc.GetMainPart().Content.Paragraphs[0].Style != "Heading" {
        t.Errorf("Expected Heading style, got %s", doc.GetMainPart().Content.Paragraphs[0].Style)
    }

    // 普通段落应该保持原样
    if doc.GetMainPart().Content.Paragraphs[1].Style != "" {
        t.Errorf("Expected empty style, got %s", doc.GetMainPart().Content.Paragraphs[1].Style)
    }
}

func TestGetQualityReport(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{
            Paragraphs: []types.Paragraph{
                {Text: "测试段落"},
            },
        },
    })

    manager := word.NewDocumentQualityManager(doc)

    // 设置一些测试指标
    manager.Metrics.ProcessingTime = 100 * time.Millisecond
    manager.Metrics.LastUpdated = time.Now()
    manager.Metrics.TotalElements = 5
    manager.Metrics.QualityScore = 0.85
    manager.Metrics.IssuesFound = 2
    manager.Metrics.IssuesFixed = 1
    manager.Metrics.MetadataScore = 0.9
    manager.Metrics.ContentScore = 0.8
    manager.Metrics.StructureScore = 0.85
    manager.Metrics.FormatScore = 0.9
    manager.Metrics.AccessibilityScore = 0.75

    report := manager.GetQualityReport()

    if report == "" {
        t.Fatal("Quality report should not be empty")
    }

    // 检查报告是否包含关键信息
    if !strings.Contains(report, "文档质量报告") {
        t.Error("Report should contain title")
    }

    if !strings.Contains(report, "处理时间") {
        t.Error("Report should contain processing time")
    }

    if !strings.Contains(report, "质量评分") {
        t.Error("Report should contain quality score")
    }
}

func TestSetAndGetQualitySettings(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{},
    })

    manager := word.NewDocumentQualityManager(doc)

    // 测试获取设置
    settings := manager.GetQualitySettings()
    if settings == nil {
        t.Fatal("Settings should not be nil")
    }

    // 测试设置新配置
    newSettings := &word.QualitySettings{
        EnableMetadataManagement: false,
        EnableContentQuality:     false,
        TextNormalization:        false,
    }

    manager.SetQualitySettings(newSettings)

    // 验证设置是否被更新
    updatedSettings := manager.GetQualitySettings()
    if updatedSettings.EnableMetadataManagement {
        t.Error("EnableMetadataManagement should be false")
    }

    if updatedSettings.EnableContentQuality {
        t.Error("EnableContentQuality should be false")
    }

    if updatedSettings.TextNormalization {
        t.Error("TextNormalization should be false")
    }
}

func TestGetQualityMetrics(t *testing.T) {
    doc := &word.Document{}
    doc.SetMainPart(&word.MainDocumentPart{
        Content: &types.DocumentContent{},
    })

    manager := word.NewDocumentQualityManager(doc)

    metrics := manager.GetQualityMetrics()
    if metrics == nil {
        t.Fatal("Metrics should not be nil")
    }

    // 检查默认指标
    if metrics.LastUpdated.IsZero() {
        t.Error("LastUpdated should be set")
    }
}
