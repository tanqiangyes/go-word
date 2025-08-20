// Package wordprocessingml provides WordprocessingML document processing functionality
package word

import (
    "fmt"
    "strings"

    "github.com/tanqiangyes/go-word/pkg/types"
)

// DocumentMerge represents a document merge operation
type DocumentMerge struct {
    SourceDocuments []*Document
    TargetDocument  *Document
    MergeOptions    MergeOptions
}

// MergeOptions defines options for document merging
type MergeOptions struct {
    MergeMode          MergeMode
    ConflictResolution ConflictResolution
    PreserveFormatting bool
    AddPageBreaks      bool
    IncludeTables      bool
    IncludeImages      bool
    MergeStyles        bool
}

// MergeMode defines how documents should be merged
type MergeMode int

const (
    // AppendMode appends content to the end
    AppendMode MergeMode = iota
    // InsertMode inserts content at specific positions
    InsertMode
    // ReplaceMode replaces existing content
    ReplaceMode
    // SelectiveMode merges only selected content
    SelectiveMode
)

// ConflictResolution defines how to handle conflicts
type ConflictResolution int

const (
    // KeepFirst keeps the first occurrence
    KeepFirst ConflictResolution = iota
    // KeepLast keeps the last occurrence
    KeepLast
    // MergeBoth merges both versions
    MergeBoth
    // SkipConflict skips conflicting content
    SkipConflict
)

// NewDocumentMerge creates a new document merge operation
func NewDocumentMerge(targetDoc *Document) *DocumentMerge {
    return &DocumentMerge{
        SourceDocuments: make([]*Document, 0),
        TargetDocument:  targetDoc,
        MergeOptions: MergeOptions{
            MergeMode:          AppendMode,
            ConflictResolution: KeepFirst,
            PreserveFormatting: true,
            AddPageBreaks:      false,
            IncludeTables:      true,
            IncludeImages:      true,
            MergeStyles:        true,
        },
    }
}

// AddSourceDocument adds a source document for merging
func (dm *DocumentMerge) AddSourceDocument(doc *Document) {
    dm.SourceDocuments = append(dm.SourceDocuments, doc)
}

// SetMergeOptions sets the merge options
func (dm *DocumentMerge) SetMergeOptions(options MergeOptions) {
    dm.MergeOptions = options
}

// MergeDocuments merges all source documents into the target document
func (dm *DocumentMerge) MergeDocuments() error {
    if len(dm.SourceDocuments) == 0 {
        return fmt.Errorf("no source documents to merge")
    }

    if dm.TargetDocument == nil {
        return fmt.Errorf("target document is nil")
    }

    switch dm.MergeOptions.MergeMode {
    case AppendMode:
        return dm.mergeAppend()
    case InsertMode:
        return dm.mergeInsert()
    case ReplaceMode:
        return dm.mergeReplace()
    case SelectiveMode:
        return dm.mergeSelective()
    default:
        return fmt.Errorf("unsupported merge mode")
    }
}

// mergeAppend appends content to the end of the target document
func (dm *DocumentMerge) mergeAppend() error {
    targetContent := dm.TargetDocument.mainPart.Content
    if targetContent == nil {
        return fmt.Errorf("target document content is nil")
    }

    for _, sourceDoc := range dm.SourceDocuments {
        if sourceDoc.mainPart == nil || sourceDoc.mainPart.Content == nil {
            continue
        }

        sourceContent := sourceDoc.mainPart.Content

        // 添加分页符（如果需要）
        if dm.MergeOptions.AddPageBreaks && len(targetContent.Paragraphs) > 0 {
            pageBreakParagraph := types.Paragraph{
                Text:  "",
                Style: "PageBreak",
                Runs:  []types.Run{},
            }
            targetContent.Paragraphs = append(targetContent.Paragraphs, pageBreakParagraph)
        }

        // 合并段落
        for _, paragraph := range sourceContent.Paragraphs {
            mergedParagraph := dm.mergeParagraph(paragraph)
            targetContent.Paragraphs = append(targetContent.Paragraphs, mergedParagraph)
        }

        // 合并表格
        if dm.MergeOptions.IncludeTables {
            for _, table := range sourceContent.Tables {
                targetContent.Tables = append(targetContent.Tables, table)
            }
        }

        // 更新文档文本
        targetContent.Text += "\n" + sourceContent.Text
    }

    return nil
}

// mergeInsert inserts content at specific positions
func (dm *DocumentMerge) mergeInsert() error {
    // 这个功能需要更复杂的实现，暂时返回错误
    return fmt.Errorf("insert mode not yet implemented")
}

// mergeReplace replaces existing content
func (dm *DocumentMerge) mergeReplace() error {
    // 这个功能需要更复杂的实现，暂时返回错误
    return fmt.Errorf("replace mode not yet implemented")
}

// mergeSelective merges only selected content
func (dm *DocumentMerge) mergeSelective() error {
    // 这个功能需要更复杂的实现，暂时返回错误
    return fmt.Errorf("selective mode not yet implemented")
}

// mergeParagraph merges a paragraph with conflict resolution
func (dm *DocumentMerge) mergeParagraph(sourceParagraph types.Paragraph) types.Paragraph {
    mergedParagraph := types.Paragraph{
        Text:  sourceParagraph.Text,
        Style: sourceParagraph.Style,
        Runs:  make([]types.Run, 0, len(sourceParagraph.Runs)),
    }

    // 合并运行
    for _, run := range sourceParagraph.Runs {
        mergedRun := dm.mergeRun(run)
        mergedParagraph.Runs = append(mergedParagraph.Runs, mergedRun)
    }

    return mergedParagraph
}

// mergeRun merges a run with formatting preservation
func (dm *DocumentMerge) mergeRun(sourceRun types.Run) types.Run {
    mergedRun := types.Run{
        Text:      sourceRun.Text,
        Bold:      sourceRun.Bold,
        Italic:    sourceRun.Italic,
        Underline: sourceRun.Underline,
        FontSize:  sourceRun.FontSize,
        FontName:  sourceRun.FontName,
    }

    // 如果保留格式，使用源格式
    if dm.MergeOptions.PreserveFormatting {
        return mergedRun
    }

    // 否则使用目标文档的默认格式
    mergedRun.FontSize = 12
    mergedRun.FontName = "Arial"
    mergedRun.Bold = false
    mergedRun.Italic = false
    mergedRun.Underline = false

    return mergedRun
}

// MergeBySection merges documents by sections
func (dm *DocumentMerge) MergeBySection(sectionMapping map[string]string) error {
    if len(dm.SourceDocuments) == 0 {
        return fmt.Errorf("no source documents to merge")
    }

    for _, targetSection := range sectionMapping {
        // 查找源文档
        var sourceDocument *Document
        for _, doc := range dm.SourceDocuments {
            if doc.mainPart != nil && doc.mainPart.Content != nil {
                // 这里需要更复杂的文档标识逻辑
                sourceDocument = doc
                break
            }
        }

        if sourceDocument == nil {
            continue
        }

        // 获取目标段落
        targetStructure, err := dm.TargetDocument.ReorganizeDocument()
        if err != nil {
            continue
        }

        targetSection, err := targetStructure.FindSectionByTitle(targetSection)
        if err != nil {
            continue
        }

        // 合并内容到目标段落
        if err := dm.mergeContentToSection(sourceDocument, targetSection); err != nil {
            return fmt.Errorf("failed to merge content to section: %w", err)
        }
    }

    return nil
}

// mergeContentToSection merges content into a specific section
func (dm *DocumentMerge) mergeContentToSection(sourceDoc *Document, targetSection *DocumentSection) error {
    if sourceDoc.mainPart == nil || sourceDoc.mainPart.Content == nil {
        return fmt.Errorf("source document content is nil")
    }

    // 获取源文档的内容
    sourceContent := sourceDoc.mainPart.Content

    // 将源内容添加到目标段落
    for _, paragraph := range sourceContent.Paragraphs {
        mergedParagraph := dm.mergeParagraph(paragraph)
        targetSection.Paragraphs = append(targetSection.Paragraphs, mergedParagraph)
    }

    return nil
}

// MergeWithConflictResolution merges documents with conflict resolution
func (dm *DocumentMerge) MergeWithConflictResolution(conflictMap map[string]ConflictResolution) error {
    // 设置冲突解决策略
    for conflictType, resolution := range conflictMap {
        switch conflictType {
        case "style":
            dm.MergeOptions.ConflictResolution = resolution
        case "content":
            // 处理内容冲突
        case "formatting":
            // 处理格式冲突
        }
    }

    return dm.MergeDocuments()
}

// GetMergeSummary returns a summary of the merge operation
func (dm *DocumentMerge) GetMergeSummary() string {
    var summary strings.Builder
    summary.WriteString("文档合并摘要:\n")
    summary.WriteString(fmt.Sprintf("源文档数量: %d\n", len(dm.SourceDocuments)))
    summary.WriteString(fmt.Sprintf("合并模式: %v\n", dm.MergeOptions.MergeMode))
    summary.WriteString(fmt.Sprintf("冲突解决: %v\n", dm.MergeOptions.ConflictResolution))
    summary.WriteString(fmt.Sprintf("保留格式: %v\n", dm.MergeOptions.PreserveFormatting))
    summary.WriteString(fmt.Sprintf("包含表格: %v\n", dm.MergeOptions.IncludeTables))
    summary.WriteString(fmt.Sprintf("包含图片: %v\n", dm.MergeOptions.IncludeImages))

    return summary.String()
}

// ValidateMerge validates the merge operation
func (dm *DocumentMerge) ValidateMerge() error {
    if len(dm.SourceDocuments) == 0 {
        return fmt.Errorf("no source documents provided")
    }

    if dm.TargetDocument == nil {
        return fmt.Errorf("target document is nil")
    }

    // 检查源文档的有效性
    for i, doc := range dm.SourceDocuments {
        if doc == nil {
            return fmt.Errorf("source document %d is nil", i)
        }
        if doc.mainPart == nil {
            return fmt.Errorf("source document %d has no main part", i)
        }
    }

    return nil
}
