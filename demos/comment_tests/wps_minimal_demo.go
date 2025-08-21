package main

import (
	"fmt"
	"log"
	"os"
	"archive/zip"
)

func main() {
	fmt.Println("🔧 开始最基础 WPS 兼容性测试...")

	// 创建最基础的 DOCX 文件
	filename := "wps_minimal_test.docx"
	
	// 创建 ZIP 文件
	zipFile, err := os.Create(filename)
	if err != nil {
		log.Fatal("Failed to create file:", err)
	}
	defer zipFile.Close()

	zipWriter := zip.NewWriter(zipFile)
	defer zipWriter.Close()

	// 添加 [Content_Types].xml
	contentTypes := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`

	contentTypesFile, err := zipWriter.Create("[Content_Types].xml")
	if err != nil {
		log.Fatal("Failed to create content types:", err)
	}
	contentTypesFile.Write([]byte(contentTypes))

	// 添加 _rels/.rels
	rels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

	relsFile, err := zipWriter.Create("_rels/.rels")
	if err != nil {
		log.Fatal("Failed to create rels:", err)
	}
	relsFile.Write([]byte(rels))

	// 添加 word/document.xml
	document := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>最基础 WPS 测试</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>这是一个测试段落。</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`

	documentFile, err := zipWriter.Create("word/document.xml")
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}
	documentFile.Write([]byte(document))

	// 添加 word/styles.xml
	styles := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`

	stylesFile, err := zipWriter.Create("word/styles.xml")
	if err != nil {
		log.Fatal("Failed to create styles:", err)
	}
	stylesFile.Write([]byte(styles))

	fmt.Printf("\n🎉 最基础 WPS 兼容性测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 最简单的文档结构")
	fmt.Println("2. 最少的命名空间")
	fmt.Println("3. 最基本的样式")
	
	fmt.Println("\n🔧 测试目的：")
	fmt.Println("- 验证最基础的 DOCX 结构是否能在 WPS 中打开")
	fmt.Println("- 如果这个文档能打开，说明问题在于我们的代码")
	fmt.Println("- 如果这个文档不能打开，说明 WPS 有严重的兼容性问题")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 WPS 中打开文档")
	fmt.Println("- 检查是否能正常显示文本内容")
	
	fmt.Println("\n🏆 这是最基础的测试，如果连这个都打不开，说明 WPS 兼容性有问题！")
}
