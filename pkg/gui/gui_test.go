package gui

import (
	"testing"
)

// TestNewGUI 测试创建GUI实例
func TestNewGUI(t *testing.T) {
	gui := NewGUI()
	if gui == nil {
		t.Fatal("GUI实例创建失败")
	}
	
	if gui.App == nil {
		t.Error("GUI应用未初始化")
	}
	
	if gui.MainWindow == nil {
		t.Error("主窗口未初始化")
	}
	
	if gui.TextArea == nil {
		t.Error("文本区域未初始化")
	}
	
	if gui.StatusBar == nil {
		t.Error("状态栏未初始化")
	}
}

// TestGUIMethods 测试GUI方法
func TestGUIMethods(t *testing.T) {
	gui := NewGUI()
	
	// 测试新建文档
	gui.newDocument()
	if gui.Document != nil {
		t.Error("新建文档后，Document应该为nil")
	}
	
	if gui.DocumentPath != "" {
		t.Error("新建文档后，DocumentPath应该为空")
	}
	
	if gui.TextArea.Text != "" {
		t.Error("新建文档后，文本区域应该为空")
	}
	
	// 测试状态栏更新
	expectedStatus := "新建文档"
	if gui.StatusBar.Text != expectedStatus {
		t.Errorf("状态栏文本不匹配，期望: %s, 实际: %s", expectedStatus, gui.StatusBar.Text)
	}
}

// TestGUIClose 测试GUI关闭
func TestGUIClose(t *testing.T) {
	gui := NewGUI()
	
	// 测试关闭方法（这里只是测试方法调用，不会真正关闭应用）
	gui.Close()
	
	// 验证GUI实例仍然存在
	if gui == nil {
		t.Error("GUI实例不应该为nil")
	}
}
