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
	
	// 创建表格
	table := system.CreateAdvancedTable("Test Table", 3, 4)
	if table == nil {
		t.Fatal("Expected table to be created")
	}
	
	if table == nil {
		t.Fatal("Expected table to be created")
	}
	
	if table.Properties.Width != 100.0 {
		t.Errorf("Expected table width 100.0, got %f", table.Properties.Width)
	}
	
	if table.Properties.Alignment != "left" {
		t.Errorf("Expected table alignment 'left', got '%s'", table.Properties.Alignment)
	}
	
	if len(table.Rows) != 3 {
		t.Errorf("Expected 3 rows, got %d", len(table.Rows))
	}
	
	if len(table.Rows[0].Cells) != 4 {
		t.Errorf("Expected 4 cells per row, got %d", len(table.Rows[0].Cells))
	}
}

