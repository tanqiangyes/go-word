package main

import (
	"fmt"
	"log"
	"os"
	"archive/zip"
)

func main() {
	fmt.Println("🔧 开始 WPS 批注功能兼容性测试...")

	// 创建批注测试的 DOCX 文件
	filename := "wps_comment_test.docx"
	
	// 创建 ZIP 文件
	zipFile, err := os.Create(filename)
	if err != nil {
		log.Fatal("Failed to create file:", err)
	}
	defer zipFile.Close()

	zipWriter := zip.NewWriter(zipFile)
	defer zipWriter.Close()

	// 添加 [Content_Types].xml - 包含批注类型
	contentTypes := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`

	contentTypesFile, err := zipWriter.Create("[Content_Types].xml")
	if err != nil {
		log.Fatal("Failed to create content types:", err)
	}
	contentTypesFile.Write([]byte(contentTypes))

	// 添加 _rels/.rels - 包含批注关系
	rels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="word/comments.xml"/>
</Relationships>`

	relsFile, err := zipWriter.Create("_rels/.rels")
	if err != nil {
		log.Fatal("Failed to create rels:", err)
	}
	relsFile.Write([]byte(rels))

	// 添加 word/_rels/document.xml.rels
	documentRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`

	documentRelsFile, err := zipWriter.Create("word/_rels/document.xml.rels")
	if err != nil {
		log.Fatal("Failed to create document rels:", err)
	}
	documentRelsFile.Write([]byte(documentRels))

	// 添加 word/document.xml - 包含批注引用
	document := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r>
        <w:t>WPS 批注功能测试</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r>
        <w:t>这是第一个段落。</w:t>
      </w:r>
      <w:commentRangeStart w:id="0"/>
      <w:commentRangeEnd w:id="0"/>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:commentReference w:id="0"/>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r>
        <w:t>这是第二个段落。</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`

	documentFile, err := zipWriter.Create("word/document.xml")
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}
	documentFile.Write([]byte(document))

	// 添加 word/styles.xml - 包含批注样式
	styles := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="character" w:styleId="CommentReference">
    <w:name w:val="Comment Reference"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed"/>
    <w:rPr>
      <w:sz w:val="16"/>
      <w:szCs w:val="16"/>
    </w:rPr>
  </w:style>
</w:styles>`

	stylesFile, err := zipWriter.Create("word/styles.xml")
	if err != nil {
		log.Fatal("Failed to create styles:", err)
	}
	stylesFile.Write([]byte(styles))

	// 添加 word/comments.xml - 批注内容
	comments := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="测试员" w:date="2024-01-01T00:00:00Z" w:initials="测试">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="CommentText"/>
      </w:pPr>
      <w:r>
        <w:t>这是一个测试批注。</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>`

	commentsFile, err := zipWriter.Create("word/comments.xml")
	if err != nil {
		log.Fatal("Failed to create comments:", err)
	}
	commentsFile.Write([]byte(comments))

	fmt.Printf("\n🎉 WPS 批注功能兼容性测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 基本的文档结构")
	fmt.Println("2. 批注引用和范围标记")
	fmt.Println("3. 批注样式定义")
	fmt.Println("4. 批注内容 XML")
	
	fmt.Println("\n🔧 测试目的：")
	fmt.Println("- 验证添加批注功能后是否还能在 WPS 中打开")
	fmt.Println("- 如果这个文档能打开，说明批注功能本身没问题")
	fmt.Println("- 如果这个文档不能打开，说明问题在于批注的 XML 结构")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 WPS 中打开文档")
	fmt.Println("- 检查是否能正常显示文本内容")
	fmt.Println("- 检查批注是否显示（如果支持的话）")
	
	fmt.Println("\n🏆 这是批注功能测试，帮助我们找到批注问题的根源！")
}
