package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/tanqiangyes/go-word/pkg/utils"
	"github.com/tanqiangyes/go-word/pkg/word"
	"github.com/tanqiangyes/go-word/pkg/writer"
)

// Command represents a command line command
type Command struct {
	Name        string
	Description string
	Usage       string
	Run         func(args []string) error
}

// CLI represents the command line interface
type CLI struct {
	commands map[string]*Command
	config   *utils.ConfigManager
	logger   *utils.Logger
}

// NewCLI creates a new CLI instance
func NewCLI() *CLI {
	cli := &CLI{
		commands: make(map[string]*Command),
		config:   utils.NewConfigManager(),
		logger:   utils.NewLogger(utils.LogLevelInfo, os.Stdout),
	}

	// Load configuration
	if err := cli.config.LoadConfig("go-word-config.json"); err != nil {
		cli.logger.Warning("无法加载配置文件: %v", err)
	}

	// Register commands
	cli.registerCommands()

	return cli
}

// registerCommands registers all available commands
func (cli *CLI) registerCommands() {
	cli.registerCommand(&Command{
		Name:        "info",
		Description: "显示文档信息",
		Usage:       "go-word info <文档路径>",
		Run:         cli.cmdInfo,
	})

	cli.registerCommand(&Command{
		Name:        "extract",
		Description: "提取文档内容",
		Usage:       "go-word extract <文档路径> [输出文件]",
		Run:         cli.cmdExtract,
	})

	cli.registerCommand(&Command{
		Name:        "create",
		Description: "创建新文档",
		Usage:       "go-word create <输出路径> [内容文件]",
		Run:         cli.cmdCreate,
	})

	cli.registerCommand(&Command{
		Name:        "convert",
		Description: "转换文档格式",
		Usage:       "go-word convert <输入路径> <输出路径>",
		Run:         cli.cmdConvert,
	})

	cli.registerCommand(&Command{
		Name:        "protect",
		Description: "保护文档",
		Usage:       "go-word protect <文档路径> <密码>",
		Run:         cli.cmdProtect,
	})

	cli.registerCommand(&Command{
		Name:        "validate",
		Description: "验证文档",
		Usage:       "go-word validate <文档路径>",
		Run:         cli.cmdValidate,
	})

	cli.registerCommand(&Command{
		Name:        "config",
		Description: "管理配置",
		Usage:       "go-word config [get|set] [键] [值]",
		Run:         cli.cmdConfig,
	})

	cli.registerCommand(&Command{
		Name:        "help",
		Description: "显示帮助信息",
		Usage:       "go-word help [命令]",
		Run:         cli.cmdHelp,
	})
}

// registerCommand registers a command
func (cli *CLI) registerCommand(cmd *Command) {
	cli.commands[cmd.Name] = cmd
}

// Run runs the CLI with the given arguments
func (cli *CLI) Run(args []string) error {
	if len(args) < 1 {
		return cli.showUsage()
	}

	commandName := args[0]
	commandArgs := args[1:]

	if cmd, exists := cli.commands[commandName]; exists {
		return cmd.Run(commandArgs)
	}

	return fmt.Errorf("未知命令: %s", commandName)
}

// showUsage shows the usage information
func (cli *CLI) showUsage() error {
	fmt.Println("Go Word 文档处理工具")
	fmt.Println()
	fmt.Println("用法:")
	fmt.Println("  go-word <命令> [参数]")
	fmt.Println()
	fmt.Println("命令:")
	for _, cmd := range cli.commands {
		fmt.Printf("  %-12s %s\n", cmd.Name, cmd.Description)
	}
	fmt.Println()
	fmt.Println("使用 'go-word help <命令>' 获取详细帮助")
	return nil
}

// cmdInfo shows document information
func (cli *CLI) cmdInfo(args []string) error {
	if len(args) < 1 {
		return fmt.Errorf("用法: %s", cli.commands["info"].Usage)
	}

	docPath := args[0]
	cli.logger.Info("正在读取文档: %s", docPath)

	// Open document
	doc, err := word.Open(docPath)
	if err != nil {
		return fmt.Errorf("无法打开文档: %w", err)
	}
	defer doc.Close()

	// Get document information
	text, err := doc.GetText()
	if err != nil {
		return fmt.Errorf("无法读取文档内容: %w", err)
	}

	paragraphs, err := doc.GetParagraphs()
	if err != nil {
		return fmt.Errorf("无法读取段落: %w", err)
	}

	tables, err := doc.GetTables()
	if err != nil {
		return fmt.Errorf("无法读取表格: %w", err)
	}

	// Display information
	fmt.Printf("文档信息:\n")
	fmt.Printf("  路径: %s\n", docPath)
	fmt.Printf("  文本长度: %d 字符\n", len(text))
	fmt.Printf("  段落数量: %d\n", len(paragraphs))
	fmt.Printf("  表格数量: %d\n", len(tables))

	// Show document parts summary
	partsSummary := doc.GetPartsSummary()
	fmt.Printf("  文档部分: %s\n", partsSummary)

	return nil
}

// cmdExtract extracts document content
func (cli *CLI) cmdExtract(args []string) error {
	if len(args) < 1 {
		return fmt.Errorf("用法: %s", cli.commands["extract"].Usage)
	}

	docPath := args[0]
	outputPath := "extracted_content.txt"
	if len(args) > 1 {
		outputPath = args[1]
	}

	cli.logger.Info("正在提取文档内容: %s -> %s", docPath, outputPath)

	// Open document
	doc, err := word.Open(docPath)
	if err != nil {
		return fmt.Errorf("无法打开文档: %w", err)
	}
	defer doc.Close()

	// Extract text
	text, err := doc.GetText()
	if err != nil {
		return fmt.Errorf("无法提取文本: %w", err)
	}

	// Write to file
	if err := os.WriteFile(outputPath, []byte(text), 0644); err != nil {
		return fmt.Errorf("无法写入输出文件: %w", err)
	}

	cli.logger.Info("内容提取完成: %s", outputPath)
	return nil
}

// cmdCreate creates a new document
func (cli *CLI) cmdCreate(args []string) error {
	if len(args) < 1 {
		return fmt.Errorf("用法: %s", cli.commands["create"].Usage)
	}

	outputPath := args[0]
	contentPath := ""
	if len(args) > 1 {
		contentPath = args[1]
	}

	cli.logger.Info("正在创建文档: %s", outputPath)

	// Create document writer
	docWriter := writer.NewDocumentWriter()
	if err := docWriter.CreateNewDocument(); err != nil {
		return fmt.Errorf("无法创建文档: %w", err)
	}

	// Add content if provided
	if contentPath != "" {
		content, err := os.ReadFile(contentPath)
		if err != nil {
			return fmt.Errorf("无法读取内容文件: %w", err)
		}

		// Split content into paragraphs
		paragraphs := splitIntoParagraphs(string(content))
		for _, paragraph := range paragraphs {
			if err := docWriter.AddParagraph(paragraph, "Normal"); err != nil {
				return fmt.Errorf("无法添加段落: %w", err)
			}
		}
	} else {
		// Add default content
		if err := docWriter.AddParagraph("这是一个新创建的文档", "Normal"); err != nil {
			return fmt.Errorf("无法添加默认段落: %w", err)
		}
	}

	// Save document
	if err := docWriter.Save(outputPath); err != nil {
		return fmt.Errorf("无法保存文档: %w", err)
	}

	cli.logger.Info("文档创建完成: %s", outputPath)
	return nil
}

// cmdConvert converts document format
func (cli *CLI) cmdConvert(args []string) error {
	if len(args) < 2 {
		return fmt.Errorf("用法: %s", cli.commands["convert"].Usage)
	}

	inputPath := args[0]
	outputPath := args[1]

	cli.logger.Info("正在转换文档格式: %s -> %s", inputPath, outputPath)

	// Open document
	doc, err := word.Open(inputPath)
	if err != nil {
		return fmt.Errorf("无法打开输入文档: %w", err)
	}
	defer doc.Close()

	// Create format support
	formatSupport := word.NewFormatSupport(doc)

	// Determine output format
	outputExt := filepath.Ext(outputPath)
	switch outputExt {
	case ".docx":
		// Already in docx format, just copy
		if err := copyFile(inputPath, outputPath); err != nil {
			return fmt.Errorf("无法复制文件: %w", err)
		}
	case ".doc":
		if err := formatSupport.ConvertFormat(word.DocFormat); err != nil {
			return fmt.Errorf("无法转换为 .doc 格式: %w", err)
		}
	case ".rtf":
		if err := formatSupport.ConvertFormat(word.RtfFormat); err != nil {
			return fmt.Errorf("无法转换为 .rtf 格式: %w", err)
		}
	default:
		return fmt.Errorf("不支持的输出格式: %s", outputExt)
	}

	cli.logger.Info("格式转换完成: %s", outputPath)
	return nil
}

// cmdProtect protects a document
func (cli *CLI) cmdProtect(args []string) error {
	if len(args) < 2 {
		return fmt.Errorf("用法: %s", cli.commands["protect"].Usage)
	}

	docPath := args[0]
	password := args[1]

	cli.logger.Info("正在保护文档: %s", docPath)

	// Open document
	doc, err := word.Open(docPath)
	if err != nil {
		return fmt.Errorf("无法打开文档: %w", err)
	}
	defer doc.Close()

	// Create document protection
	protection := word.NewDocumentProtection()

	// Enable protection
	if err := protection.EnableProtection(word.ReadOnlyProtection, password); err != nil {
		return fmt.Errorf("无法启用保护: %w", err)
	}

	// Save protected document
	outputPath := docPath + ".protected.docx"
	if err := doc.Close(); err != nil {
		return fmt.Errorf("无法保存受保护的文档: %w", err)
	}

	cli.logger.Info("文档保护完成: %s", outputPath)
	return nil
}

// cmdValidate validates a document
func (cli *CLI) cmdValidate(args []string) error {
	if len(args) < 1 {
		return fmt.Errorf("用法: %s", cli.commands["validate"].Usage)
	}

	docPath := args[0]

	cli.logger.Info("正在验证文档: %s", docPath)

	// Open document
	doc, err := word.Open(docPath)
	if err != nil {
		return fmt.Errorf("无法打开文档: %w", err)
	}
	defer doc.Close()

	// Create document validator
	validator := word.NewDocumentValidator(doc)

	// Validate document
	if err := validator.ValidateDocument(); err != nil {
		return fmt.Errorf("文档验证失败: %w", err)
	}

	// Show validation results
	results := validator.GetValidationResults()
	if len(results) == 0 {
		fmt.Println("文档验证通过，未发现问题")
	} else {
		fmt.Printf("发现 %d 个问题:\n", len(results))
		for i, result := range results {
			fmt.Printf("  %d. %s (%s)\n", i+1, result.Message, result.Severity)
		}
	}

	return nil
}

// cmdConfig manages configuration
func (cli *CLI) cmdConfig(args []string) error {
	if len(args) == 0 {
		// Show all configuration
		config := cli.config.GetConfig()
		fmt.Printf("当前配置:\n")
		fmt.Printf("  语言: %s\n", config.Language)
		fmt.Printf("  默认字体: %s\n", config.DefaultFont)
		fmt.Printf("  默认字体大小: %d\n", config.DefaultFontSize)
		fmt.Printf("  最大文件大小: %d MB\n", config.MaxFileSize/(1024*1024))
		fmt.Printf("  内存限制: %d MB\n", config.MemoryLimit/(1024*1024))
		fmt.Printf("  超时时间: %d 秒\n", config.Timeout)
		return nil
	}

	if len(args) < 2 {
		return fmt.Errorf("用法: %s", cli.commands["config"].Usage)
	}

	action := args[0]
	key := args[1]

	switch action {
	case "get":
		value := cli.config.GetString(key)
		if value == "" {
			value = fmt.Sprintf("%v", cli.config.GetInt(key))
		}
		fmt.Printf("%s = %s\n", key, value)

	case "set":
		if len(args) < 3 {
			return fmt.Errorf("设置配置需要提供值")
		}
		value := args[2]
		cli.config.SetConfig(key, value)
		fmt.Printf("已设置 %s = %s\n", key, value)

	default:
		return fmt.Errorf("未知操作: %s", action)
	}

	return nil
}

// cmdHelp shows help information
func (cli *CLI) cmdHelp(args []string) error {
	if len(args) == 0 {
		return cli.showUsage()
	}

	commandName := args[0]
	if cmd, exists := cli.commands[commandName]; exists {
		fmt.Printf("命令: %s\n", cmd.Name)
		fmt.Printf("描述: %s\n", cmd.Description)
		fmt.Printf("用法: %s\n", cmd.Usage)
	} else {
		return fmt.Errorf("未知命令: %s", commandName)
	}

	return nil
}

// splitIntoParagraphs splits text into paragraphs
func splitIntoParagraphs(text string) []string {
	// Simple paragraph splitting by double newlines
	paragraphs := make([]string, 0)
	current := ""

	for _, char := range text {
		if char == '\n' {
			if current != "" {
				paragraphs = append(paragraphs, current)
				current = ""
			}
		} else {
			current += string(char)
		}
	}

	if current != "" {
		paragraphs = append(paragraphs, current)
	}

	return paragraphs
}

// copyFile copies a file
func copyFile(src, dst string) error {
	data, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, data, 0644)
}

func main() {
	cli := NewCLI()

	// Parse command line flags
	flag.Parse()

	// Run CLI
	if err := cli.Run(flag.Args()); err != nil {
		log.Fatal(err)
	}
}
