package main

import (
	"fmt"
	"log"
	"os"
	"archive/zip"
)

func main() {
	fmt.Println("🔧 开始 WPS 简化 DocumentWriter 兼容性测试...")

	// 创建简化测试的 DOCX 文件
	filename := "wps_simplified_writer_test.docx"
	
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
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
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

	// 添加 word/_rels/document.xml.rels
	documentRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`

	documentRelsFile, err := zipWriter.Create("word/_rels/document.xml.rels")
	if err != nil {
		log.Fatal("Failed to create document rels:", err)
	}
	documentRelsFile.Write([]byte(documentRels))

	// 添加 word/document.xml - 模拟我们的代码生成的文档
	document := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Times New Roman"/>
          <w:sz w:val="22"/>
          <w:szCs w:val="22"/>
          <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
        </w:rPr>
        <w:t>WPS 简化 DocumentWriter 测试</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:widowControl w:val="0"/>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Times New Roman"/>
          <w:sz w:val="22"/>
          <w:szCs w:val="22"/>
          <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
        </w:rPr>
        <w:t>这是第一个测试段落。</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:widowControl w:val="0"/>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Times New Roman"/>
          <w:sz w:val="22"/>
          <w:szCs w:val="22"/>
          <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
        </w:rPr>
        <w:t>这是第二个测试段落。</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`

	documentFile, err := zipWriter.Create("word/document.xml")
	if err != nil {
		log.Fatal("Failed to create document:", err)
	}
	documentFile.Write([]byte(document))

	// 添加 word/styles.xml - 模拟我们的代码生成的样式
	styles := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Times New Roman"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:widowControl w:val="0"/>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:widowControl w:val="0"/>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
</w:styles>`

	stylesFile, err := zipWriter.Create("word/styles.xml")
	if err != nil {
		log.Fatal("Failed to create styles:", err)
	}
	stylesFile.Write([]byte(styles))

	// 添加 word/fontTable.xml
	fontTable := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E0002AFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="宋体">
    <w:panose1 w:val="00000000000000000000"/>
    <w:charset w:val="86"/>
    <w:family w:val="auto"/>
    <w:pitch w:val="default"/>
    <w:sig w:usb0="00000003" w:usb1="288F0000" w:usb2="00000016" w:usb3="00000000" w:csb0="00040001" w:csb1="00000000"/>
  </w:font>
</w:fonts>`

	fontTableFile, err := zipWriter.Create("word/fontTable.xml")
	if err != nil {
		log.Fatal("Failed to create font table:", err)
	}
	fontTableFile.Write([]byte(fontTable))

	// 添加 word/settings.xml
	settings := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat/>
  <w:rsids>
    <w:rsidRoot w:val="00000000"/>
  </w:rsids>
</w:settings>`

	settingsFile, err := zipWriter.Create("word/settings.xml")
	if err != nil {
		log.Fatal("Failed to create settings:", err)
	}
	settingsFile.Write([]byte(settings))

	fmt.Printf("\n🎉 WPS 简化 DocumentWriter 兼容性测试完成！文件已保存: %s\n", filename)
	fmt.Println("\n📋 测试内容：")
	fmt.Println("1. 模拟我们的代码生成的文档结构")
	fmt.Println("2. 包含复杂的命名空间和属性")
	fmt.Println("3. 包含字体表和设置文件")
	
	fmt.Println("\n🔧 测试目的：")
	fmt.Println("- 验证我们代码生成的文档结构是否能在 WPS 中打开")
	fmt.Println("- 如果这个文档能打开，说明问题在于其他部分")
	fmt.Println("- 如果这个文档不能打开，说明问题在于我们的 XML 结构")
	
	fmt.Println("\n🔍 验证要点：")
	fmt.Println("- 在 WPS 中打开文档")
	fmt.Println("- 检查是否能正常显示文本内容")
	fmt.Println("- 检查字体和样式是否正常")
	
	fmt.Println("\n🏆 这是模拟我们代码的测试，帮助我们找到真正的问题！")
}
