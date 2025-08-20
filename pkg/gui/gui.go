package gui

import (
	"fmt"
	"path/filepath"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/tanqiangyes/go-word/pkg/word"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

// GUI 图形用户界面
type GUI struct {
	app          fyne.App
	mainWindow   fyne.Window
	document     *word.Document
	documentPath string
	textArea     *widget.Entry
	statusBar    *widget.Label
}

// NewGUI 创建新的GUI实例
func NewGUI() *GUI {
	gui := &GUI{
		app: app.New(),
	}
	gui.setupMainWindow()
	return gui
}

// setupMainWindow 设置主窗口
func (gui *GUI) setupMainWindow() {
	gui.mainWindow = gui.app.NewWindow("Go Word - 文档处理器")
	gui.mainWindow.Resize(fyne.NewSize(1000, 700))

	// 创建菜单栏
	gui.createMenuBar()

	// 创建工具栏
	toolbar := gui.createToolbar()

	// 创建文本区域
	gui.textArea = widget.NewMultiLineEntry()
	gui.textArea.SetPlaceHolder("在这里输入或编辑文档内容...")

	// 创建状态栏
	gui.statusBar = widget.NewLabel("就绪")

	// 创建主布局
	content := container.NewBorder(toolbar, gui.statusBar, nil, nil, gui.textArea)
	gui.mainWindow.SetContent(content)
}

// createMenuBar 创建菜单栏
func (gui *GUI) createMenuBar() {
	// 文件菜单
	fileMenu := fyne.NewMenu("文件",
		fyne.NewMenuItem("新建", gui.newDocument),
		fyne.NewMenuItem("打开", gui.openDocument),
		fyne.NewMenuItem("保存", gui.saveDocument),
		fyne.NewMenuItem("另存为", gui.saveDocumentAs),
		fyne.NewMenuItem("导出为PDF", gui.exportToPDF),
		fyne.NewMenuItemSeparator(),
		fyne.NewMenuItem("退出", func() {
			gui.app.Quit()
		}),
	)

	// 编辑菜单
	editMenu := fyne.NewMenu("编辑",
		fyne.NewMenuItem("撤销", func() {}),
		fyne.NewMenuItem("重做", func() {}),
		fyne.NewMenuItemSeparator(),
		fyne.NewMenuItem("剪切", func() {}),
		fyne.NewMenuItem("复制", func() {}),
		fyne.NewMenuItem("粘贴", func() {}),
	)

	// 格式菜单
	formatMenu := fyne.NewMenu("格式",
		fyne.NewMenuItem("字体", gui.showFontDialog),
		fyne.NewMenuItem("段落", gui.showParagraphDialog),
		fyne.NewMenuItem("样式", gui.showStyleDialog),
	)

	// 工具菜单
	toolsMenu := fyne.NewMenu("工具",
		fyne.NewMenuItem("图片处理", gui.showImageProcessor),
		fyne.NewMenuItem("图表生成", gui.showChartGenerator),
		fyne.NewMenuItem("文件嵌入", gui.showFileEmbedder),
		fyne.NewMenuItem("性能优化", gui.showPerformanceOptimizer),
	)

	// 帮助菜单
	helpMenu := fyne.NewMenu("帮助",
		fyne.NewMenuItem("关于", gui.showAbout),
		fyne.NewMenuItem("用户指南", gui.showUserGuide),
	)

	mainMenu := fyne.NewMainMenu(fileMenu, editMenu, formatMenu, toolsMenu, helpMenu)
	gui.mainWindow.SetMainMenu(mainMenu)
}

// createToolbar 创建工具栏
func (gui *GUI) createToolbar() *widget.Toolbar {
	toolbar := widget.NewToolbar(
		widget.NewToolbarAction(storage.NewFolderOpenIcon(), gui.openDocument),
		widget.NewToolbarAction(storage.NewDocumentSaveIcon(), gui.saveDocument),
		widget.NewToolbarSeparator(),
		widget.NewToolbarAction(storage.NewMediaSkipPreviousIcon(), func() {}), // 撤销
		widget.NewToolbarAction(storage.NewMediaSkipNextIcon(), func() {}),     // 重做
		widget.NewToolbarSeparator(),
		widget.NewToolbarAction(storage.NewMediaPlayIcon(), func() {}),  // 粗体
		widget.NewToolbarAction(storage.NewMediaPauseIcon(), func() {}), // 斜体
		widget.NewToolbarAction(storage.NewMediaStopIcon(), func() {}),  // 下划线
		widget.NewToolbarSeparator(),
		widget.NewToolbarAction(storage.NewMediaFastForwardIcon(), gui.exportToPDF), // 导出PDF
	)
	return toolbar
}

// newDocument 新建文档
func (gui *GUI) newDocument() {
	gui.document = nil
	gui.documentPath = ""
	gui.textArea.SetText("")
	gui.statusBar.SetText("新建文档")
	gui.mainWindow.SetTitle("Go Word - 文档处理器")
}

// openDocument 打开文档
func (gui *GUI) openDocument() {
	dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
		if err != nil {
			dialog.ShowError(err, gui.mainWindow)
			return
		}
		if reader == nil {
			return
		}
		defer reader.Close()

		filePath := reader.URI().Path()
		if !strings.HasSuffix(strings.ToLower(filePath), ".docx") {
			dialog.ShowError(fmt.Errorf("只支持.docx格式文件"), gui.mainWindow)
			return
		}

		// 打开文档
		doc, err := word.Open(filePath)
		if err != nil {
			dialog.ShowError(fmt.Errorf("打开文档失败: %w", err), gui.mainWindow)
			return
		}

		// 获取文档内容
		text, err := doc.GetText()
		if err != nil {
			dialog.ShowError(fmt.Errorf("读取文档内容失败: %w", err), gui.mainWindow)
			return
		}

		gui.document = doc
		gui.documentPath = filePath
		gui.textArea.SetText(text)
		gui.statusBar.SetText(fmt.Sprintf("已打开: %s", filepath.Base(filePath)))
		gui.mainWindow.SetTitle(fmt.Sprintf("Go Word - %s", filepath.Base(filePath)))
	}, gui.mainWindow)
}

// saveDocument 保存文档
func (gui *GUI) saveDocument() {
	if gui.document == nil {
		gui.saveDocumentAs()
		return
	}

	// 获取文本内容
	text := gui.textArea.Text

	// 创建文档写入器
	w := writer.NewDocumentWriter()

	// 添加段落
	if err := w.AddParagraph(text, "Normal"); err != nil {
		dialog.ShowError(fmt.Errorf("保存文档失败: %w", err), gui.mainWindow)
		return
	}

	// 保存文档
	if err := w.Save(gui.documentPath); err != nil {
		dialog.ShowError(fmt.Errorf("保存文档失败: %w", err), gui.mainWindow)
		return
	}

	gui.statusBar.SetText("文档已保存")
}

// saveDocumentAs 另存为
func (gui *GUI) saveDocumentAs() {
	dialog.ShowFileSave(func(writer fyne.URIWriteCloser, err error) {
		if err != nil {
			dialog.ShowError(err, gui.mainWindow)
			return
		}
		if writer == nil {
			return
		}
		defer writer.Close()

		filePath := writer.URI().Path()
		if !strings.HasSuffix(strings.ToLower(filePath), ".docx") {
			filePath += ".docx"
		}

		// 获取文本内容
		text := gui.textArea.Text

		// 创建文档写入器
		w := writer.NewDocumentWriter()

		// 添加段落
		if err := w.AddParagraph(text, "Normal"); err != nil {
			dialog.ShowError(fmt.Errorf("保存文档失败: %w", err), gui.mainWindow)
			return
		}

		// 保存文档
		if err := w.Save(filePath); err != nil {
			dialog.ShowError(fmt.Errorf("保存文档失败: %w", err), gui.mainWindow)
			return
		}

		gui.documentPath = filePath
		gui.statusBar.SetText(fmt.Sprintf("文档已保存为: %s", filepath.Base(filePath)))
		gui.mainWindow.SetTitle(fmt.Sprintf("Go Word - %s", filepath.Base(filePath)))
	}, gui.mainWindow)
}

// exportToPDF 导出为PDF
func (gui *GUI) exportToPDF() {
	if gui.document == nil {
		dialog.ShowError(fmt.Errorf("请先打开一个文档"), gui.mainWindow)
		return
	}

	dialog.ShowFileSave(func(writer fyne.URIWriteCloser, err error) {
		if err != nil {
			dialog.ShowError(err, gui.mainWindow)
			return
		}
		if writer == nil {
			return
		}
		defer writer.Close()

		filePath := writer.URI().Path()
		if !strings.HasSuffix(strings.ToLower(filePath), ".pdf") {
			filePath += ".pdf"
		}

		// 创建PDF导出器
		pdfExporter := word.NewPDFExporter(gui.document, nil)

		// 导出PDF
		if err := pdfExporter.ExportToPDF(nil, filePath); err != nil {
			dialog.ShowError(fmt.Errorf("PDF导出失败: %w", err), gui.mainWindow)
			return
		}

		gui.statusBar.SetText(fmt.Sprintf("PDF已导出: %s", filepath.Base(filePath)))
	}, gui.mainWindow)
}

// showFontDialog 显示字体对话框
func (gui *GUI) showFontDialog() {
	dialog.ShowInformation("字体设置", "字体设置功能正在开发中...", gui.mainWindow)
}

// showParagraphDialog 显示段落对话框
func (gui *GUI) showParagraphDialog() {
	dialog.ShowInformation("段落设置", "段落设置功能正在开发中...", gui.mainWindow)
}

// showStyleDialog 显示样式对话框
func (gui *GUI) showStyleDialog() {
	dialog.ShowInformation("样式设置", "样式设置功能正在开发中...", gui.mainWindow)
}

// showImageProcessor 显示图片处理器
func (gui *GUI) showImageProcessor() {
	dialog.ShowInformation("图片处理器", "图片处理器功能正在开发中...", gui.mainWindow)
}

// showChartGenerator 显示图表生成器
func (gui *GUI) showChartGenerator() {
	dialog.ShowInformation("图表生成器", "图表生成器功能正在开发中...", gui.mainWindow)
}

// showFileEmbedder 显示文件嵌入器
func (gui *GUI) showFileEmbedder() {
	dialog.ShowInformation("文件嵌入器", "文件嵌入器功能正在开发中...", gui.mainWindow)
}

// showPerformanceOptimizer 显示性能优化器
func (gui *GUI) showPerformanceOptimizer() {
	dialog.ShowInformation("性能优化器", "性能优化器功能正在开发中...", gui.mainWindow)
}

// showAbout 显示关于对话框
func (gui *GUI) showAbout() {
	dialog.ShowInformation("关于", "Go Word 文档处理器\n版本: 1.0.0\n\n一个用Go语言重写的Microsoft Open XML SDK，专门用于Word文档处理。", gui.mainWindow)
}

// showUserGuide 显示用户指南
func (gui *GUI) showUserGuide() {
	dialog.ShowInformation("用户指南", "用户指南功能正在开发中...\n\n请参考项目文档获取详细使用说明。", gui.mainWindow)
}

// Run 运行GUI应用
func (gui *GUI) Run() {
	gui.mainWindow.ShowAndRun()
}

// Close 关闭GUI应用
func (gui *GUI) Close() {
	gui.app.Quit()
}
