# Go OpenXML SDK

一个用 Go 语言重写的 Microsoft Open XML SDK，专门用于 Word 文档处理。

## 特性

- 🚀 **高性能**：优化的解析速度，最小化内存占用
- 📄 **完整支持**：支持 Word 文档格式（.docx），包含新老版本
- 🔧 **Go 风格**：遵循 Go 语言最佳实践和惯用法
- 📚 **详细文档**：完整的 API 文档和使用示例
- 🛡️ **错误处理**：详细的错误信息和日志记录

## 快速开始

### 安装

```bash
go get github.com/go-word
```

### 基本使用

```go
package main

import (
    "fmt"
    "log"
    
    "github.com/go-word/pkg/wordprocessingml"
)

func main() {
    // 打开 Word 文档
    doc, err := wordprocessingml.Open("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()
    
    // 获取文档内容
    content, err := doc.GetText()
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Println("文档内容:", content)
}
```

## 项目结构

```
go-word/
├── pkg/
│   ├── opc/              # OPC 容器处理
│   ├── wordprocessingml/ # Word 文档处理
│   ├── parser/           # 解析器
│   ├── writer/           # 写入器
│   └── utils/            # 工具函数
├── examples/             # 使用示例
├── docs/                # 文档
└── tests/               # 测试
```

## 开发状态

- [x] 项目初始化
- [ ] OPC 容器基础功能
- [ ] WordprocessingML 解析
- [ ] 文档内容提取
- [ ] 样式和格式解析
- [ ] 文档修改功能
- [ ] 高级操作功能

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可证

MIT License 