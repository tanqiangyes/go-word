// Package wordprocessingml provides WordprocessingML document processing functionality
package wordprocessingml

import (
	"fmt"
	"strings"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// AdvancedTableSystem represents advanced table functionality
type AdvancedTableSystem struct {
	// 表格管理器
	TableManager *TableManager
	
	// 表格样式
	TableStyles map[string]*AdvancedTableStyle
	
	// 表格模板
	TableTemplates map[string]*TableTemplate
	
	// 表格验证器
	TableValidator *TableValidator
}

// TableManager manages advanced table operations
type TableManager struct {
	// 表格集合
	Tables map[string]*AdvancedTable
	
	// 表格操作历史
	History []TableOperation
	
	// 表格统计
	Statistics *TableStatistics
}

// AdvancedTable represents an advanced table
type AdvancedTable struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 表格结构
	Rows        []AdvancedTableRow
	Columns     []AdvancedTableColumn
	Headers     []AdvancedTableHeader
	
	// 表格属性
	Properties  *AdvancedTableProperties
	
	// 表格样式
	Style       *AdvancedTableStyle
	
	// 表格数据
	Data        [][]string
	
	// 其他属性
	Visible     bool
	Locked      bool
}

// AdvancedTableRow represents an advanced table row
type AdvancedTableRow struct {
	// 基础信息
	Index       int
	ID          string
	
	// 行属性
	Height      float64
	MinHeight   float64
	MaxHeight   float64
	Hidden      bool
	RepeatHeader bool
	
	// 行样式
	Style       *RowStyle
	
	// 单元格
	Cells       []AdvancedTableCell
	
	// 其他属性
	AllowBreak  bool
	KeepTogether bool
}

// AdvancedTableColumn represents an advanced table column
type AdvancedTableColumn struct {
	// 基础信息
	Index       int
	ID          string
	
	// 列属性
	Width       float64
	MinWidth    float64
	MaxWidth    float64
	Hidden      bool
	
	// 列样式
	Style       *ColumnStyle
	
	// 其他属性
	AutoFit     bool
	BestFit     bool
}

// AdvancedTableHeader represents an advanced table header
type AdvancedTableHeader struct {
	// 基础信息
	ID          string
	Title       string
	
	// 标题属性
	Level       int
	Style       string
	Alignment   string
	
	// 其他属性
	Visible     bool
	Repeat      bool
}

// AdvancedTableProperties represents advanced table properties
type AdvancedTableProperties struct {
	// 基础属性
	Width       float64
	Height      float64
	Alignment   string
	Position    string
	
	// 边框属性
	Borders     *AdvancedTableBorders
	
	// 背景属性
	Background  *AdvancedTableBackground
	
	// 布局属性
	Layout      *AdvancedTableLayout
	
	// 其他属性
	AllowOverlap bool
	AllowBreak   bool
	KeepTogether bool
}

// AdvancedTableBorders represents advanced table borders
type AdvancedTableBorders struct {
	// 外边框
	Top         *BorderStyle
	Bottom      *BorderStyle
	Left        *BorderStyle
	Right       *BorderStyle
	
	// 内边框
	InsideH     *BorderStyle
	InsideV     *BorderStyle
	
	// 其他属性
	Shadow      bool
	ShadowColor string
	ShadowSize  float64
}

// BorderStyle represents a border style
type BorderStyle struct {
	// 基础属性
	Style       string
	Size        float64
	Color       string
	Space       float64
	
	// 其他属性
	Visible     bool
	Shadow      bool
}

// AdvancedTableBackground represents advanced table background
type AdvancedTableBackground struct {
	// 背景颜色
	Color       string
	Opacity     float64
	
	// 背景图片
	Image       []byte
	ImageWidth  float64
	ImageHeight float64
	
	// 其他属性
	Pattern     string
	PatternColor string
}

// AdvancedTableLayout represents advanced table layout
type AdvancedTableLayout struct {
	// 布局类型
	Type        TableLayoutType
	
	// 布局属性
	Width       float64
	Alignment   string
	Indent      float64
	
	// 其他属性
	FixedLayout bool
	AutoLayout  bool
}

// TableLayoutType defines table layout type
type TableLayoutType int

const (
	// FixedLayoutType for fixed layout
	FixedLayoutType TableLayoutType = iota
	// AutoLayoutType for auto layout
	AutoLayoutType
	// CustomLayoutType for custom layout
	CustomLayoutType
)

// AdvancedTableCell represents an advanced table cell
type AdvancedTableCell struct {
	// 基础信息
	RowIndex    int
	ColIndex    int
	ID          string
	
	// 单元格属性
	Width       float64
	Height      float64
	Alignment   string
	VerticalAlignment string
	
	// 合并属性
	RowSpan     int
	ColSpan     int
	
	// 单元格内容
	Content     *CellContent
	
	// 单元格样式
	Style       *CellStyle
	
	// 其他属性
	Hidden      bool
	Locked      bool
}

// CellContent 使用 advanced_formatting.go 中的定义

// CellImage represents a cell image
type CellImage struct {
	// 基础信息
	ID          string
	Data        []byte
	Format      string
	
	// 图片属性
	Width       float64
	Height      float64
	Alignment   string
	
	// 其他属性
	Caption     string
	AltText     string
}

// CellObject represents a cell object
type CellObject struct {
	// 基础信息
	ID          string
	Type        string
	Data        []byte
	
	// 对象属性
	Width       float64
	Height      float64
	Alignment   string
	
	// 其他属性
	Visible     bool
	Locked      bool
}

// CellStyle represents cell style
type CellStyle struct {
	// 背景样式
	Background  *CellBackground
	
	// 边框样式
	Borders     *CellBorders
	
	// 字体样式
	Font        *CellFont
	
	// 其他样式
	Padding     *CellPadding
	Margin      *CellMargin
}

// CellBackground represents cell background
type CellBackground struct {
	// 背景颜色
	Color       string
	Opacity     float64
	
	// 背景图片
	Image       []byte
	ImageWidth  float64
	ImageHeight float64
	
	// 其他属性
	Pattern     string
	PatternColor string
}

// CellBorders 使用 advanced_formatting.go 中的定义

// CellFont represents cell font
type CellFont struct {
	// 字体属性
	Name        string
	Size        float64
	Bold        bool
	Italic      bool
	Underline   bool
	Strike      bool
	
	// 颜色属性
	Color       string
	Highlight   string
	
	// 其他属性
	Superscript bool
	Subscript   bool
	SmallCaps   bool
}

// CellPadding represents cell padding
type CellPadding struct {
	// 内边距
	Top         float64
	Bottom      float64
	Left        float64
	Right       float64
}

// CellMargin represents cell margin
type CellMargin struct {
	// 外边距
	Top         float64
	Bottom      float64
	Left        float64
	Right       float64
}

// RowStyle represents row style
type RowStyle struct {
	// 行属性
	Height      float64
	MinHeight   float64
	MaxHeight   float64
	
	// 行样式
	Background  *CellBackground
	Borders     *CellBorders
	
	// 其他属性
	Hidden      bool
	RepeatHeader bool
}

// ColumnStyle represents column style
type ColumnStyle struct {
	// 列属性
	Width       float64
	MinWidth    float64
	MaxWidth    float64
	
	// 列样式
	Background  *CellBackground
	Borders     *CellBorders
	
	// 其他属性
	Hidden      bool
	AutoFit     bool
	BestFit     bool
}

// AdvancedTableStyle represents advanced table style
type AdvancedTableStyle struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 样式属性
	Properties  *AdvancedTableProperties
	
	// 样式继承
	BasedOn     string
	Next        string
	
	// 其他属性
	Hidden      bool
	Locked      bool
}

// TableTemplate represents a table template
type TableTemplate struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 模板结构
	Rows        int
	Columns     int
	Headers     []string
	
	// 模板样式
	Style       *AdvancedTableStyle
	
	// 模板数据
	SampleData  [][]string
	
	// 其他属性
	Category    string
	Tags        []string
}

// TableValidator validates table structure and content
type TableValidator struct {
	// 验证规则
	Rules       []TableValidationRule
	
	// 验证结果
	Results     []TableValidationResult
	
	// 验证设置
	Settings    *ValidationSettings
}

// TableValidationRule represents a table validation rule
type TableValidationRule struct {
	// 基础信息
	ID          string
	Name        string
	Description string
	
	// 规则类型
	Type        ValidationRuleType
	
	// 规则条件
	Condition   string
	Severity    ValidationSeverity
	
	// 其他属性
	Enabled     bool
	Priority    int
}

// ValidationRuleType defines validation rule type
type ValidationRuleType int

const (
	// StructureRule for structure validation
	StructureRule ValidationRuleType = iota
	// ContentRule for content validation
	ContentRule
	// StyleRule for style validation
	StyleRule
	// FormatRule for format validation
	FormatRule
)

// ValidationSeverity defines validation severity
type ValidationSeverity int

const (
	// InfoSeverity for info level
	InfoSeverity ValidationSeverity = iota
	// WarningSeverity for warning level
	WarningSeverity
	// ErrorSeverity for error level
	ErrorSeverity
	// CriticalSeverity for critical level
	CriticalSeverity
)

// TableValidationResult represents a table validation result
type TableValidationResult struct {
	// 基础信息
	ID          string
	RuleID      string
	TableID     string
	
	// 验证结果
	Valid       bool
	Severity    ValidationSeverity
	Message     string
	
	// 位置信息
	Row         int
	Column      int
	CellID      string
	
	// 其他属性
	Timestamp   string
	Fixed       bool
}

// ValidationSettings represents validation settings
type ValidationSettings struct {
	// 验证选项
	ValidateStructure bool
	ValidateContent   bool
	ValidateStyle     bool
	ValidateFormat    bool
	
	// 自动修复
	AutoFix          bool
	AutoFixLevel     ValidationSeverity
	
	// 其他设置
	StopOnError      bool
	MaxErrors        int
}

// TableOperation represents a table operation
type TableOperation struct {
	// 基础信息
	ID          string
	Type        TableOperationType
	TableID     string
	
	// 操作详情
	Description string
	Parameters  map[string]interface{}
	
	// 时间信息
	Timestamp   string
	Duration    float64
	
	// 其他属性
	Success     bool
	Error       string
}

// TableOperationType defines table operation type
type TableOperationType int

const (
	// CreateOperation for create
	CreateOperation TableOperationType = iota
	// UpdateOperation for update
	UpdateOperation
	// DeleteOperation for delete
	DeleteOperation
	// MergeOperation for merge
	MergeOperation
	// SplitOperation for split
	SplitOperation
	// SortOperation for sort
	SortOperation
	// FilterOperation for filter
	FilterOperation
)

// TableStatistics represents table statistics
type TableStatistics struct {
	// 基础统计
	TotalTables int
	TotalRows    int
	TotalColumns int
	TotalCells   int
	
	// 样式统计
	StyledTables int
	CustomStyles int
	Templates    int
	
	// 验证统计
	ValidTables  int
	InvalidTables int
	Errors       int
	Warnings     int
}

// NewAdvancedTableSystem creates new advanced table system
func NewAdvancedTableSystem() *AdvancedTableSystem {
	return &AdvancedTableSystem{
		TableManager: &TableManager{
			Tables: make(map[string]*AdvancedTable),
			History: make([]TableOperation, 0),
			Statistics: &TableStatistics{},
		},
		TableStyles: make(map[string]*AdvancedTableStyle),
		TableTemplates: make(map[string]*TableTemplate),
		TableValidator: &TableValidator{
			Rules: make([]TableValidationRule, 0),
			Results: make([]TableValidationResult, 0),
			Settings: &ValidationSettings{
				ValidateStructure: true,
				ValidateContent:   true,
				ValidateStyle:     true,
				ValidateFormat:    true,
				AutoFix:          false,
				AutoFixLevel:     WarningSeverity,
				StopOnError:      false,
				MaxErrors:        100,
			},
		},
	}
}

// CreateAdvancedTable creates a new advanced table
func (ats *AdvancedTableSystem) CreateAdvancedTable(name string, rows, cols int) *AdvancedTable {
	table := &AdvancedTable{
		ID:          fmt.Sprintf("table_%d", len(ats.TableManager.Tables)+1),
		Name:        name,
		Description: fmt.Sprintf("Advanced table with %d rows and %d columns", rows, cols),
		Rows:        make([]AdvancedTableRow, rows),
		Columns:     make([]AdvancedTableColumn, cols),
		Headers:     make([]AdvancedTableHeader, cols),
		Properties:  &AdvancedTableProperties{
			Width:       100.0,
			Height:      0.0,
			Alignment:   "left",
			Position:    "inline",
			Borders:     &AdvancedTableBorders{},
			Background:  &AdvancedTableBackground{},
			Layout:      &AdvancedTableLayout{
				Type:        AutoLayoutType,
				Width:       100.0,
				Alignment:   "left",
				Indent:      0.0,
				FixedLayout: false,
				AutoLayout:  true,
			},
			AllowOverlap: false,
			AllowBreak:   true,
			KeepTogether: false,
		},
		Style:       nil,
		Data:        make([][]string, rows),
		Visible:     true,
		Locked:      false,
	}
	
	// 初始化行
	for i := 0; i < rows; i++ {
		table.Rows[i] = AdvancedTableRow{
			Index:       i,
			ID:          fmt.Sprintf("row_%d", i),
			Height:      20.0,
			MinHeight:   10.0,
			MaxHeight:   100.0,
			Hidden:      false,
			RepeatHeader: false,
			Style:       &RowStyle{},
			Cells:       make([]AdvancedTableCell, cols),
			AllowBreak:  true,
			KeepTogether: false,
		}
		
		// 初始化数据行
		table.Data[i] = make([]string, cols)
		
		// 初始化单元格
		for j := 0; j < cols; j++ {
			table.Rows[i].Cells[j] = AdvancedTableCell{
				RowIndex:    i,
				ColIndex:    j,
				ID:          fmt.Sprintf("cell_%d_%d", i, j),
				Width:       20.0,
				Height:      20.0,
				Alignment:   "left",
				VerticalAlignment: "top",
				RowSpan:     1,
				ColSpan:     1,
				Content:     &CellContent{
					Text:       fmt.Sprintf("单元格 %d-%d", i+1, j+1),
					Paragraphs: make([]types.Paragraph, 0),
					Images:     make([]Image, 0),
				},
				Style:       &CellStyle{},
				Hidden:      false,
				Locked:      false,
			}
			
			// 设置默认数据
			table.Data[i][j] = fmt.Sprintf("单元格 %d-%d", i+1, j+1)
		}
	}
	
	// 初始化列
	for j := 0; j < cols; j++ {
		table.Columns[j] = AdvancedTableColumn{
			Index:       j,
			ID:          fmt.Sprintf("col_%d", j),
			Width:       20.0,
			MinWidth:    10.0,
			MaxWidth:    100.0,
			Hidden:      false,
			Style:       &ColumnStyle{},
			AutoFit:     true,
			BestFit:     true,
		}
		
		// 初始化表头
		table.Headers[j] = AdvancedTableHeader{
			ID:          fmt.Sprintf("header_%d", j),
			Title:       fmt.Sprintf("列 %d", j+1),
			Level:       1,
			Style:       "Header",
			Alignment:   "center",
			Visible:     true,
			Repeat:      true,
		}
	}
	
	// 添加到管理器
	ats.TableManager.Tables[table.ID] = table
	
	// 记录操作
	operation := TableOperation{
		ID:          fmt.Sprintf("op_%d", len(ats.TableManager.History)+1),
		Type:        CreateOperation,
		TableID:     table.ID,
		Description: fmt.Sprintf("Created table '%s' with %d rows and %d columns", name, rows, cols),
		Parameters: map[string]interface{}{
			"name": name,
			"rows": rows,
			"cols": cols,
		},
		Timestamp:   "now",
		Duration:    0.0,
		Success:     true,
		Error:       "",
	}
	ats.TableManager.History = append(ats.TableManager.History, operation)
	
	// 更新统计
	ats.TableManager.Statistics.TotalTables++
	ats.TableManager.Statistics.TotalRows += rows
	ats.TableManager.Statistics.TotalColumns += cols
	ats.TableManager.Statistics.TotalCells += rows * cols
	
	return table
}

// MergeCells merges cells in a table
func (ats *AdvancedTableSystem) MergeCells(tableID string, startRow, startCol, endRow, endCol int) error {
	table := ats.TableManager.Tables[tableID]
	if table == nil {
		return fmt.Errorf("table not found: %s", tableID)
	}
	
	if startRow < 0 || startRow >= len(table.Rows) ||
		endRow < 0 || endRow >= len(table.Rows) ||
		startCol < 0 || startCol >= len(table.Columns) ||
		endCol < 0 || endCol >= len(table.Columns) {
		return fmt.Errorf("invalid cell range")
	}
	
	if startRow > endRow || startCol > endCol {
		return fmt.Errorf("invalid range order")
	}
	
	// 计算合并范围
	rowSpan := endRow - startRow + 1
	colSpan := endCol - startCol + 1
	
	// 设置主单元格
	mainCell := &table.Rows[startRow].Cells[startCol]
	mainCell.RowSpan = rowSpan
	mainCell.ColSpan = colSpan
	
	// 合并内容
	mergedText := ""
	for i := startRow; i <= endRow; i++ {
		for j := startCol; j <= endCol; j++ {
			if i == startRow && j == startCol {
				mergedText += table.Data[i][j]
			} else {
				mergedText += " " + table.Data[i][j]
				// 隐藏被合并的单元格
				table.Rows[i].Cells[j].Hidden = true
			}
		}
	}
	
	mainCell.Content.Text = mergedText
	table.Data[startRow][startCol] = mergedText
	
	// 记录操作
	operation := TableOperation{
		ID:          fmt.Sprintf("op_%d", len(ats.TableManager.History)+1),
		Type:        MergeOperation,
		TableID:     tableID,
		Description: fmt.Sprintf("Merged cells from (%d,%d) to (%d,%d)", startRow, startCol, endRow, endCol),
		Parameters: map[string]interface{}{
			"startRow": startRow,
			"startCol": startCol,
			"endRow":   endRow,
			"endCol":   endCol,
			"rowSpan":  rowSpan,
			"colSpan":  colSpan,
		},
		Timestamp:   "now",
		Duration:    0.0,
		Success:     true,
		Error:       "",
	}
	ats.TableManager.History = append(ats.TableManager.History, operation)
	
	return nil
}

// SplitCells splits merged cells
func (ats *AdvancedTableSystem) SplitCells(tableID string, row, col int) error {
	table := ats.TableManager.Tables[tableID]
	if table == nil {
		return fmt.Errorf("table not found: %s", tableID)
	}
	
	if row < 0 || row >= len(table.Rows) || col < 0 || col >= len(table.Columns) {
		return fmt.Errorf("invalid cell position")
	}
	
	cell := &table.Rows[row].Cells[col]
	if cell.RowSpan == 1 && cell.ColSpan == 1 {
		return fmt.Errorf("cell is not merged")
	}
	
	// 重置合并属性
	cell.RowSpan = 1
	cell.ColSpan = 1
	
	// 显示被隐藏的单元格
	for i := row; i < row+cell.RowSpan; i++ {
		for j := col; j < col+cell.ColSpan; j++ {
			if i != row || j != col {
				table.Rows[i].Cells[j].Hidden = false
			}
		}
	}
	
	// 记录操作
	operation := TableOperation{
		ID:          fmt.Sprintf("op_%d", len(ats.TableManager.History)+1),
		Type:        SplitOperation,
		TableID:     tableID,
		Description: fmt.Sprintf("Split cell at (%d,%d)", row, col),
		Parameters: map[string]interface{}{
			"row": row,
			"col": col,
		},
		Timestamp:   "now",
		Duration:    0.0,
		Success:     true,
		Error:       "",
	}
	ats.TableManager.History = append(ats.TableManager.History, operation)
	
	return nil
}

// GetTableSummary returns a summary of all tables
func (ats *AdvancedTableSystem) GetTableSummary() string {
	var summary strings.Builder
	summary.WriteString("高级表格系统摘要:\n")
	summary.WriteString(fmt.Sprintf("表格数量: %d\n", ats.TableManager.Statistics.TotalTables))
	summary.WriteString(fmt.Sprintf("总行数: %d\n", ats.TableManager.Statistics.TotalRows))
	summary.WriteString(fmt.Sprintf("总列数: %d\n", ats.TableManager.Statistics.TotalColumns))
	summary.WriteString(fmt.Sprintf("总单元格数: %d\n", ats.TableManager.Statistics.TotalCells))
	summary.WriteString(fmt.Sprintf("样式表格: %d\n", ats.TableManager.Statistics.StyledTables))
	summary.WriteString(fmt.Sprintf("自定义样式: %d\n", ats.TableManager.Statistics.CustomStyles))
	summary.WriteString(fmt.Sprintf("模板数量: %d\n", ats.TableManager.Statistics.Templates))
	summary.WriteString(fmt.Sprintf("有效表格: %d\n", ats.TableManager.Statistics.ValidTables))
	summary.WriteString(fmt.Sprintf("无效表格: %d\n", ats.TableManager.Statistics.InvalidTables))
	summary.WriteString(fmt.Sprintf("错误数量: %d\n", ats.TableManager.Statistics.Errors))
	summary.WriteString(fmt.Sprintf("警告数量: %d\n", ats.TableManager.Statistics.Warnings))
	summary.WriteString(fmt.Sprintf("操作历史: %d\n", len(ats.TableManager.History)))
	
	return summary.String()
} 