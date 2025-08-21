package main

import (
	"fmt"
	"log"
	"os"
	"archive/zip"
)

func main() {
	fmt.Println("🔧 开始 WPS 最简化样式测试...")

	// 创建最简化测试的 DOCX 文件
	filename := "wps_minimal_style_test.docx"
	
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

	// 添加 word/document.xml - 最简化的文档
	document := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r>
        <w:t>最简化样式测试</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
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

	// 添加 word/styles.xml - 最简化的样式定义
	styles := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr>
  </w:style>
</w:styles>`

	stylesFile, err := zipWriter.Create("word/styles.xml")
	if err != nil {
		log.Fatal("Failed to create styles:", err)
	}
	stylesFile.Write([]byte(styles))

	fmt.Printf("\n🎉 WPS 最简化样式测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 最简化的文档结构")
	fmt.Println("2. 只包含 Normal 样式")
	fmt.Println("3. 没有任何复杂的样式定义")
	
	fmt.Println("\n🔧 测试目的：")
	fmt.Println("- 验证最基本的样式是否能在 Word 中正常打开")
	fmt.Println("- 如果这个文档能正常打开，说明问题在于我们的复杂样式")
	fmt.Println("- 如果这个文档还有问题，说明问题在于基础结构")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 Word 中打开文档，检查是否有错误")
	fmt.Println("- 在 WPS 中打开文档，检查是否能正常显示")
	
	fmt.Println("\n🏆 这是最基础的测试，帮助我们找到问题的根源！")
}
