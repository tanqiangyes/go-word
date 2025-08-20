package main

import (
	"fmt"

	"github.com/tanqiangyes/go-word/pkg/word"
)

func DemoDocumentProtection() {
	fmt.Println("=== Go Word 文档保护功能演示 ===")

	// 演示文档保护功能
	demoDocumentProtection()

	fmt.Println("文档保护功能演示完成！")
}

func demoDocumentProtection() {
	fmt.Println("1. 创建文档保护系统...")

	// 创建文档保护系统
	protection := word.NewDocumentProtection()

	fmt.Println("2. 演示文档保护启用...")

	// 启用只读保护
	if err := protection.EnableProtection(word.ReadOnlyProtection, "password123"); err != nil {
		fmt.Printf("启用只读保护失败: %v\n", err)
	} else {
		fmt.Println("✓ 启用只读保护成功")
	}

	// 检查保护状态
	fmt.Printf("保护状态: %v\n", protection.Settings.Enabled)
	fmt.Printf("保护类型: %v\n", protection.Settings.ProtectionType)
	fmt.Printf("密码保护: %v\n", protection.Settings.Password != "")

	fmt.Println("3. 演示权限管理...")

	// 添加用户权限
	userPermissions := map[string]bool{
		"read":   true,
		"edit":   false,
		"delete": false,
		"format": false,
		"print":  true,
		"share":  false,
	}

	if err := protection.AddUserPermission("user1", "张三", "zhangsan@example.com", userPermissions); err != nil {
		fmt.Printf("添加用户权限失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加用户权限成功")
	}

	// 检查权限
	fmt.Printf("用户1读取权限: %v\n", protection.CheckPermission("user1", "read"))
	fmt.Printf("用户1编辑权限: %v\n", protection.CheckPermission("user1", "edit"))
	fmt.Printf("用户1删除权限: %v\n", protection.CheckPermission("user1", "delete"))
	fmt.Printf("用户1打印权限: %v\n", protection.CheckPermission("user1", "print"))

	// 添加另一个用户权限
	adminPermissions := map[string]bool{
		"read":   true,
		"edit":   true,
		"delete": true,
		"format": true,
		"print":  true,
		"share":  true,
	}

	if err := protection.AddUserPermission("admin1", "管理员", "admin@example.com", adminPermissions); err != nil {
		fmt.Printf("添加管理员权限失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加管理员权限成功")
	}

	// 检查管理员权限
	fmt.Printf("管理员读取权限: %v\n", protection.CheckPermission("admin1", "read"))
	fmt.Printf("管理员编辑权限: %v\n", protection.CheckPermission("admin1", "edit"))
	fmt.Printf("管理员删除权限: %v\n", protection.CheckPermission("admin1", "delete"))
	fmt.Printf("管理员打印权限: %v\n", protection.CheckPermission("admin1", "print"))

	fmt.Println("4. 演示水印功能...")

	// 添加文本水印
	if err := protection.AddWatermark("机密文档", "机密", word.TextWatermark); err != nil {
		fmt.Printf("添加文本水印失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加文本水印成功")
	}

	// 添加Logo水印
	if err := protection.AddWatermark("公司Logo", "公司名称", word.LogoWatermark); err != nil {
		fmt.Printf("添加Logo水印失败: %v\n", err)
	} else {
		fmt.Println("✓ 添加Logo水印成功")
	}

	// 显示水印信息
	fmt.Printf("水印数量: %d\n", len(protection.Watermark.Watermarks))
	for i, watermark := range protection.Watermark.Watermarks {
		fmt.Printf("水印%d: %s (%s)\n", i+1, watermark.Name, watermark.Text)
	}

	fmt.Println("5. 演示保护类型切换...")

	// 切换到注释保护
	if err := protection.EnableProtection(word.CommentsProtection, "password123"); err != nil {
		fmt.Printf("切换到注释保护失败: %v\n", err)
	} else {
		fmt.Println("✓ 切换到注释保护成功")
	}

	fmt.Printf("新保护类型: %v\n", protection.Settings.ProtectionType)
	fmt.Printf("限制为注释: %v\n", protection.Settings.EditRestrictions.RestrictToComments)

	// 切换到跟踪更改保护
	if err := protection.EnableProtection(word.TrackChangesProtection, "password123"); err != nil {
		fmt.Printf("切换到跟踪更改保护失败: %v\n", err)
	} else {
		fmt.Println("✓ 切换到跟踪更改保护成功")
	}

	fmt.Printf("新保护类型: %v\n", protection.Settings.ProtectionType)
	fmt.Printf("限制为跟踪更改: %v\n", protection.Settings.EditRestrictions.RestrictToTrackChanges)

	fmt.Println("6. 演示保护禁用...")

	// 禁用保护
	if err := protection.DisableProtection("password123"); err != nil {
		fmt.Printf("禁用保护失败: %v\n", err)
	} else {
		fmt.Println("✓ 禁用保护成功")
	}

	fmt.Printf("保护状态: %v\n", protection.Settings.Enabled)
	fmt.Printf("保护类型: %v\n", protection.Settings.ProtectionType)

	// 检查权限（保护禁用后应该全部允许）
	fmt.Printf("用户1读取权限（保护禁用后）: %v\n", protection.CheckPermission("user1", "read"))
	fmt.Printf("用户1编辑权限（保护禁用后）: %v\n", protection.CheckPermission("user1", "edit"))

	fmt.Println("7. 演示加密功能...")

	// 启用加密
	protection.Encryption.Settings.Enabled = true
	protection.Encryption.Settings.Algorithm = word.AES256Algorithm
	protection.Encryption.Settings.KeySize = 256
	protection.Encryption.Settings.EncryptContent = true

	fmt.Printf("加密状态: %v\n", protection.Encryption.Settings.Enabled)
	fmt.Printf("加密算法: %v\n", protection.Encryption.Settings.Algorithm)
	fmt.Printf("密钥大小: %d\n", protection.Encryption.Settings.KeySize)
	fmt.Printf("加密内容: %v\n", protection.Encryption.Settings.EncryptContent)

	fmt.Println("8. 演示数字签名功能...")

	// 启用数字签名
	protection.DigitalSignature.Settings.Enabled = true
	protection.DigitalSignature.Settings.Required = false
	protection.DigitalSignature.Settings.Multiple = true
	protection.DigitalSignature.Settings.SignContent = true
	protection.DigitalSignature.Settings.Timestamp = true

	fmt.Printf("数字签名状态: %v\n", protection.DigitalSignature.Settings.Enabled)
	fmt.Printf("签名必需: %v\n", protection.DigitalSignature.Settings.Required)
	fmt.Printf("多重签名: %v\n", protection.DigitalSignature.Settings.Multiple)
	fmt.Printf("签名内容: %v\n", protection.DigitalSignature.Settings.SignContent)
	fmt.Printf("时间戳: %v\n", protection.DigitalSignature.Settings.Timestamp)

	fmt.Println("9. 显示保护摘要...")
	fmt.Println(protection.GetProtectionSummary())

	fmt.Println("文档保护功能演示完成！")
}


