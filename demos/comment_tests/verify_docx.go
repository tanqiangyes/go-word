package main

import (
	"archive/zip"
	"fmt"
	"io"
	"log"
	"os"
	"strings"
)

func main() {
	filename := "simple_test.docx"
	
	fmt.Printf("验证 DOCX 文件: %s\n", filename)
	
	// 检查文件是否存在
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		log.Fatal("文件不存在:", filename)
	}
	
	// 打开 ZIP 文件
	reader, err := zip.OpenReader(filename)
	if err != nil {
		log.Fatal("无法打开 DOCX 文件:", err)
	}
	defer reader.Close()
	
	fmt.Println("\n文件内容:")
	fmt.Println("==========")
	
	// 列出所有文件
	for _, file := range reader.File {
		fmt.Printf("- %s\n", file.Name)
		
		// 如果是 XML 文件，显示前几行内容
		if strings.HasSuffix(file.Name, ".xml") {
			rc, err := file.Open()
			if err != nil {
				fmt.Printf("  无法读取文件: %v\n", err)
				continue
			}
			
			content, err := io.ReadAll(rc)
			rc.Close()
			if err != nil {
				fmt.Printf("  读取失败: %v\n", err)
				continue
			}
			
			// 显示前 200 个字符
			contentStr := string(content)
			if len(contentStr) > 200 {
				contentStr = contentStr[:200] + "..."
			}
			fmt.Printf("  内容预览: %s\n", contentStr)
		}
	}
	
	fmt.Println("\n验证完成！")
}
