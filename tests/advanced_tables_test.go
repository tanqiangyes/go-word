package tests

import (
	"testing"
	
	"github.com/tanqiangyes/go-word/pkg/wordprocessingml"
)

func TestNewAdvancedTableSystem(t *testing.T) {
	system := wordprocessingml.NewAdvancedTableSystem()
	
	if system == nil {
		t.Fatal("Expected AdvancedTableSystem to be created")
	}
	
	if system.TableManager == nil {
		t.Error("Expected TableManager to be initialized")
	}
	
	if system.TableStyles == nil {
		t.Error("Expected TableStyles to be initialized")
	}
	
	if system.TableTemplates == nil {
		t.Error("Expected TableTemplates to be initialized")
	}
	
	if system.TableValidator == nil {
		t.Error("Expected TableValidator to be initialized")
	}
}

func TestCreateAdvancedTable(t *testing.T) {
	system := wordprocessingml.NewAdvancedTableSystem()
	
	// 创建表格属性
	properties := wordprocessingml.AdvancedTableProperties{
		Width:        100,
		Alignment:    "center",
		CellSpacing:  0,
		CellPadding:  5,
		Borders:      true,
		Shading:      false,
		Layout:       "fixed",
		Caption:      "Test Table",
		Description:  "A test table for validation",
	}
	
	// 创建表格
	table, err := system.CreateAdvancedTable(3, 4, properties)
	if err != nil {
		t.Fatalf("Failed to create advanced table: %v", err)
	}
	
	if table == nil {
		t.Fatal("Expected table to be created")
	}
	
	if table.Properties.Width != 100 {
		t.Errorf("Expected table width 100, got %d", table.Properties.Width)
	}
	
	if table.Properties.Alignment != "center" {
		t.Errorf("Expected table alignment 'center', got '%s'", table.Properties.Alignment)
	}
	
	if len(table.Rows) != 3 {
		t.Errorf("Expected 3 rows, got %d", len(table.Rows))
	}
	
	if len(table.Rows[0].Cells) != 4 {
		t.Errorf("Expected 4 cells per row, got %d", len(table.Rows[0].Cells))
	}
}

