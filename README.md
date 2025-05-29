# doc

一个用于从 Microsoft Word .doc 二进制文件中提取文本的 Go 包。

## 功能特点

- 从 Microsoft Word .doc 二进制文件中提取纯文本
- 处理压缩和非压缩文本格式
- 支持多种字符编码，包括中文字符
- 简单易用的 API 接口

## 安装

```bash
go get github.com/lee501/doc
```

## 使用方法

```go
package main

import (
	"fmt"
	"github.com/lee501/doc"
	"os"
)

func main() {
	// 打开 Word 文档
	file, err := os.Open("document.doc")
	if err != nil {
		panic(err)
	}
	defer file.Close()

	// 从文档中提取文本
	text, err := doc.ParseDoc(file)
	if err != nil {
		panic(err)
	}

	// 将读取器转换为字符串并打印
	fmt.Println(text)
}
```

## 功能详情
1. 分离压缩和非压缩文本处理
- 将 translateText 函数拆分为 translateCompressedText 和 translateUncompressedText
- 压缩文本通常使用单字节编码（如 ANSI/CP1252）
- 非压缩文本通常使用双字节 Unicode 编码

2. 增强的 Unicode 支持
- 在 translateUncompressedText 中正确处理双字节 Unicode 字符
- 使用 binary.LittleEndian.Uint16 读取 Unicode 码点
- 将 Unicode 码点转换为 UTF-8 输出

3. 改进的中文字符处理
- 添加 handleANSICharacter 函数处理可能的中文字符
- 集成 golang.org/x/text/encoding/simplifiedchinese 包以支持 GBK 编码
- 对高字节字符尝试 GBK 解码

4. 更好的字符映射
- 增强 replaceCompressed 函数，正确将 Windows-1252 特殊字符转换为 UTF-8
- 添加详细的 Unicode 码点注释

5. 编码检测
- 添加 detectChineseEncoding 辅助函数来识别可能的中文编码
- 为将来更智能的编码检测奠定基础

## 依赖项

- [mattetti/filebuffer](https://github.com/mattetti/filebuffer)
- [richardlehane/mscfb](https://github.com/richardlehane/mscfb)
- [golang.org/x/text/encoding/simplifiedchinese](https://pkg.go.dev/golang.org/x/text/encoding/simplifiedchinese)

## 许可证(LICENSE)

本项目采用 MIT 许可证 - 详情请参阅 [LICENSE](LICENSE) 文件。
