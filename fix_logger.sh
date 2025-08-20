#!/bin/bash

# 修复Logger调用的格式问题
# 将 map[string]interface{} 格式改为格式化字符串

echo "开始修复Logger调用格式..."

# 修复 advanced_styles.go
sed -i 's/ass\.logger\.Info("合并属性冲突", map\[string\]interface{}{/ass.logger.Info("合并属性冲突，样式ID: %s, 冲突数量: %d", original.ID, len(conflict.ConflictingProperties))/' advanced_styles.go
sed -i 's/ass\.logger\.Info("属性冲突合并完成", map\[string\]interface{}{/ass.logger.Info("属性冲突合并完成，样式ID: %s, 合并属性数: %d", mergedStyle.ID, len(conflict.ConflictingProperties))/' advanced_styles.go
sed -i 's/ass\.logger\.Info("合并继承冲突", map\[string\]interface{}{/ass.logger.Info("合并继承冲突，样式ID: %s, 原始继承: %s, 新继承: %s", original.ID, original.BasedOn, new.BasedOn)/' advanced_styles.go
sed -i 's/ass\.logger\.Info("继承冲突合并完成", map\[string\]interface{}{/ass.logger.Info("继承冲突合并完成，样式ID: %s, 最优继承链长度: %d", mergedStyle.ID, len(optimalChain))/' advanced_styles.go
sed -i 's/ass\.logger\.Info("合并优先级冲突", map\[string\]interface{}{/ass.logger.Info("合并优先级冲突，样式ID: %s, 原始优先级: %d, 新优先级: %d", original.ID, original.Priority, new.Priority)/' advanced_styles.go
sed -i 's/ass\.logger\.Info("选择新样式（更高优先级）", map\[string\]interface{}{/ass.logger.Info("选择新样式（更高优先级），样式ID: %s", new.ID)/' advanced_styles.go
sed -i 's/ass\.logger\.Info("保留原始样式（更高或相等优先级）", map\[string\]interface{}{/ass.logger.Info("保留原始样式（更高或相等优先级），样式ID: %s", original.ID)/' advanced_styles.go
sed -i 's/ass\.logger\.Info("合并格式冲突", map\[string\]interface{}{/ass.logger.Info("合并格式冲突，样式ID: %s, 冲突数量: %d", original.ID, len(conflict.ConflictingProperties))/' advanced_styles.go
sed -i 's/ass\.logger\.Info("格式冲突合并完成", map\[string\]interface{}{/ass.logger.Info("格式冲突合并完成，样式ID: %s, 合并属性数: %d", mergedStyle.ID, len(conflict.ConflictingProperties))/' advanced_styles.go
sed -i 's/ass\.logger\.Info("使用默认合并策略", map\[string\]interface{}{/ass.logger.Info("使用默认合并策略，样式ID: %s", original.ID)/' advanced_styles.go

# 修复 custom_ribbon.go
sed -i 's/cr\.logger\.Info("选项卡添加成功", map\[string\]interface{}{/cr.logger.Info("选项卡添加成功，选项卡ID: %s, 名称: %s", tab.ID, tab.Name)/' custom_ribbon.go
sed -i 's/cr\.logger\.Info("组添加成功", map\[string\]interface{}{/cr.logger.Info("组添加成功，组ID: %s, 名称: %s", group.ID, group.Name)/' custom_ribbon.go
sed -i 's/cr\.logger\.Info("控件添加成功", map\[string\]interface{}{/cr.logger.Info("控件添加成功，控件ID: %s, 类型: %s", control.ID, control.Type)/' custom_ribbon.go
sed -i 's/cr\.logger\.Info("回调函数注册成功", map\[string\]interface{}{/cr.logger.Info("回调函数注册成功，控件ID: %s", controlID)/' custom_ribbon.go
sed -i 's/cr\.logger\.Info("控件触发成功", map\[string\]interface{}{/cr.logger.Info("控件触发成功，控件ID: %s", controlID)/' custom_ribbon.go
sed -i 's/cr\.logger\.Info("控件更新成功", map\[string\]interface{}{/cr.logger.Info("控件更新成功，控件ID: %s", controlID)/' custom_ribbon.go
sed -i 's/cr\.logger\.Info("控件移除成功", map\[string\]interface{}{/cr.logger.Info("控件移除成功，控件ID: %s", controlID)/' custom_ribbon.go
sed -i 's/cr\.logger\.Info("功能区导入成功", map\[string\]interface{}{/cr.logger.Info("功能区导入成功，文件路径: %s", filePath)/' custom_ribbon.go

# 修复 api_design.go
sed -i 's/b\.logger\.Info("文档标题已设置", map\[string\]interface{}{/b.logger.Info("文档标题已设置: %s", title)/' api_design.go
sed -i 's/b\.logger\.Info("文档作者已设置", map\[string\]interface{}{/b.logger.Info("文档作者已设置: %s", author)/' api_design.go
sed -i 's/b\.logger\.Info("文档保护已应用", map\[string\]interface{}{/b.logger.Info("文档保护已应用，保护类型: %s", protectionType)/' api_design.go
sed -i 's/b\.logger\.Info("文档验证已应用", map\[string\]interface{}{/b.logger.Info("文档验证已应用，验证类型: %s", validationType)/' api_design.go
sed -i 's/b\.logger\.Info("评论已添加", map\[string\]interface{}{/b.logger.Info("评论已添加，评论ID: %s", comment.ID)/' api_design.go

# 修复 api_enhanced.go
sed -i 's/b\.logger\.Info("设置文档标题", map\[string\]interface{}{/b.logger.Info("设置文档标题: %s", title)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档标题已设置", map\[string\]interface{}{/b.logger.Info("文档标题已设置: %s", title)/' api_enhanced.go
sed -i 's/b\.logger\.Info("设置文档作者", map\[string\]interface{}{/b.logger.Info("设置文档作者: %s", author)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档作者已设置", map\[string\]interface{}{/b.logger.Info("文档作者已设置: %s", author)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档主题已设置", map\[string\]interface{}{/b.logger.Info("文档主题已设置: %s", subject)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档关键词已设置", map\[string\]interface{}{/b.logger.Info("文档关键词已设置: %s", keywords)/' api_enhanced.go
sed -i 's/b\.logger\.Info("应用文档保护", map\[string\]interface{}{/b.logger.Info("应用文档保护，保护类型: %s", protectionType)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档保护已应用", map\[string\]interface{}{/b.logger.Info("文档保护已应用，保护类型: %s", protectionType)/' api_enhanced.go
sed -i 's/b\.logger\.Info("应用文档验证", map\[string\]interface{}{/b.logger.Info("应用文档验证，验证类型: %s", validationType)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档验证已应用", map\[string\]interface{}{/b.logger.Info("文档验证已应用，验证类型: %s", validationType)/' api_enhanced.go
sed -i 's/b\.logger\.Info("添加段落到文档", map\[string\]interface{}{/b.logger.Info("添加段落到文档，段落ID: %s", paragraph.ID)/' api_enhanced.go
sed -i 's/b\.logger\.Info("段落已添加到文档", map\[string\]interface{}{/b.logger.Info("段落已添加到文档，段落ID: %s", paragraph.ID)/' api_enhanced.go
sed -i 's/b\.logger\.Info("添加表格到文档", map\[string\]interface{}{/b.logger.Info("添加表格到文档，表格ID: %s", table.ID)/' api_enhanced.go
sed -i 's/b\.logger\.Info("表格已添加到文档", map\[string\]interface{}{/b.logger.Info("表格已添加到文档，表格ID: %s", table.ID)/' api_enhanced.go
sed -i 's/b\.logger\.Info("添加图片到文档", map\[string\]interface{}{/b.logger.Info("添加图片到文档，图片ID: %s", image.ID)/' api_enhanced.go
sed -i 's/b\.logger\.Info("图片已添加到文档", map\[string\]interface{}{/b.logger.Info("图片已添加到文档，图片ID: %s", image.ID)/' api_enhanced.go
sed -i 's/b\.logger\.Info("保存文档", map\[string\]interface{}{/b.logger.Info("保存文档，文件路径: %s", filePath)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档已保存", map\[string\]interface{}{/b.logger.Info("文档已保存，文件路径: %s", filePath)/' api_enhanced.go
sed -i 's/b\.logger\.Info("保存文档为指定格式", map\[string\]interface{}{/b.logger.Info("保存文档为指定格式，文件路径: %s, 格式: %s", filePath, format)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档已保存为指定格式", map\[string\]interface{}{/b.logger.Info("文档已保存为指定格式，文件路径: %s, 格式: %s", filePath, format)/' api_enhanced.go
sed -i 's/b\.logger\.Info("导出文档", map\[string\]interface{}{/b.logger.Info("导出文档，文件路径: %s, 格式: %s", filePath, format)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档已导出为PDF", map\[string\]interface{}{/b.logger.Info("文档已导出为PDF，文件路径: %s", filePath)/' api_enhanced.go
sed -i 's/b\.logger\.Info("RTF导出功能待实现", map\[string\]interface{}{/b.logger.Info("RTF导出功能待实现，文件路径: %s", filePath)/' api_enhanced.go
sed -i 's/b\.logger\.Info("HTML导出功能待实现", map\[string\]interface{}{/b.logger.Info("HTML导出功能待实现，文件路径: %s", filePath)/' api_enhanced.go
sed -i 's/b\.logger\.Info("文档已导出为TXT", map\[string\]interface{}{/b.logger.Info("文档已导出为TXT，文件路径: %s", filePath)/' api_enhanced.go

# 修复 collaborative_editor.go
sed -i 's/ce\.logger\.Info("协作会话已创建", map\[string\]interface{}{/ce.logger.Info("协作会话已创建，会话ID: %s", session.ID)/' collaborative_editor.go
sed -i 's/ce\.logger\.Info("用户已加入会话", map\[string\]interface{}{/ce.logger.Info("用户已加入会话，用户ID: %s, 会话ID: %s", userID, sessionID)/' collaborative_editor.go
sed -i 's/ce\.logger\.Info("用户已离开会话", map\[string\]interface{}{/ce.logger.Info("用户已离开会话，用户ID: %s, 会话ID: %s", userID, sessionID)/' collaborative_editor.go
sed -i 's/ce\.logger\.Info("操作已应用", map\[string\]interface{}{/ce.logger.Info("操作已应用，操作ID: %s, 类型: %s", operation.ID, operation.Type)/' collaborative_editor.go
sed -i 's/ce\.logger\.Info("冲突已解决", map\[string\]interface{}{/ce.logger.Info("冲突已解决，冲突ID: %s", conflict.ID)/' collaborative_editor.go
sed -i 's/ce\.logger\.Info("检测到冲突", map\[string\]interface{}{/ce.logger.Info("检测到冲突，冲突ID: %s, 类型: %s", conflict.ID, conflict.Type)/' collaborative_editor.go
sed -i 's/ce\.logger\.Info("自动清理完成", map\[string\]interface{}{/ce.logger.Info("自动清理完成，清理会话数: %d", len(sessionsToClean))/' collaborative_editor.go

# 修复 format_support.go
sed -i 's/fs\.logger\.Info("开始转换到.doc格式", map\[string\]interface{}{/fs.logger.Info("开始转换到.doc格式，文件路径: %s", outputPath)/' format_support.go
sed -i 's/fs\.logger\.Info("开始转换到.docm格式", map\[string\]interface{}{/fs.logger.Info("开始转换到.docm格式，文件路径: %s", outputPath)/' format_support.go
sed -i 's/fs\.logger\.Info("开始转换到RTF格式", map\[string\]interface{}{/fs.logger.Info("开始转换到RTF格式，文件路径: %s", outputPath)/' format_support.go
sed -i 's/fs\.logger\.Info("成功生成.doc文件", map\[string\]interface{}{/fs.logger.Info("成功生成.doc文件，文件路径: %s", outputPath)/' format_support.go
sed -i 's/fs\.logger\.Info("成功保存RTF文件", map\[string\]interface{}{/fs.logger.Info("成功保存RTF文件，文件路径: %s", outputPath)/' format_support.go
sed -i 's/fs\.logger\.Info("创建宏启用容器", map\[string\]interface{}{/fs.logger.Info("创建宏启用容器，容器ID: %s", containerID)/' format_support.go
sed -i 's/fs\.logger\.Info("保存宏启用文档", map\[string\]interface{}{/fs.logger.Info("保存宏启用文档，文件路径: %s", outputPath)/' format_support.go

# 修复 pdf_exporter.go
sed -i 's/pe\.logger\.Info("开始PDF导出", map\[string\]interface{}{/pe.logger.Info("开始PDF导出，文件路径: %s", outputPath)/' pdf_exporter.go
sed -i 's/pe\.logger\.Info("PDF导出完成", map\[string\]interface{}{/pe.logger.Info("PDF导出完成，文件路径: %s", outputPath)/' pdf_exporter.go
sed -i 's/pe\.logger\.Info("保存PDF文件", map\[string\]interface{}{/pe.logger.Info("保存PDF文件，文件路径: %s", outputPath)/' pdf_exporter.go

# 修复 file_embedder.go
sed -i 's/fe\.logger\.Info("文件已存在，使用现有嵌入", map\[string\]interface{}{/fe.logger.Info("文件已存在，使用现有嵌入，文件ID: %s", fileID)/' file_embedder.go
sed -i 's/fe\.logger\.Info("文件嵌入成功", map\[string\]interface{}{/fe.logger.Info("文件嵌入成功，文件ID: %s, 大小: %d", fileID, fileSize)/' file_embedder.go
sed -i 's/fe\.logger\.Info("链接创建成功", map\[string\]interface{}{/fe.logger.Info("链接创建成功，链接ID: %s", linkID)/' file_embedder.go

# 修复 style_merger.go
sed -i 's/sm\.logger\.Info("开始合并样式", map\[string\]interface{}{/sm.logger.Info("开始合并样式，样式ID: %s", styleID)/' style_merger.go
sed -i 's/sm\.logger\.Info("样式合并完成", map\[string\]interface{}{/sm.logger.Info("样式合并完成，样式ID: %s", styleID)/' style_merger.go
sed -i 's/sm\.logger\.Info("字体属性合并完成", map\[string\]interface{}{/sm.logger.Info("字体属性合并完成，样式ID: %s", styleID)/' style_merger.go
sed -i 's/sm\.logger\.Info("段落属性合并完成", map\[string\]interface{}{/sm.logger.Info("段落属性合并完成，样式ID: %s", styleID)/' style_merger.go
sed -i 's/sm\.logger\.Info("表格属性合并完成", map\[string\]interface{}{/sm.logger.Info("表格属性合并完成，样式ID: %s", styleID)/' style_merger.go
sed -i 's/sm\.logger\.Info("格式属性合并完成", map\[string\]interface{}{/sm.logger.Info("格式属性合并完成，样式ID: %s", styleID)/' style_merger.go
sed -i 's/sm\.logger\.Info("样式验证完成", map\[string\]interface{}{/sm.logger.Info("样式验证完成，样式ID: %s", styleID)/' style_merger.go
sed -i 's/sm\.logger\.Info("样式合并失败", map\[string\]interface{}{/sm.logger.Info("样式合并失败，样式ID: %s, 错误: %v", styleID, err)/' style_merger.go
sed -i 's/sm\.logger\.Info("样式合并警告", map\[string\]interface{}{/sm.logger.Info("样式合并警告，样式ID: %s, 警告: %s", styleID, warning)/' style_merger.go

# 修复 template.go
sed -i 's/t\.logger\.Info("模板验证开始", map\[string\]interface{}{/t.logger.Info("模板验证开始，模板ID: %s", t.ID)/' template.go
sed -i 's/t\.logger\.Info("模板验证完成", map\[string\]interface{}{/t.logger.Info("模板验证完成，模板ID: %s", t.ID)/' template.go
sed -i 's/t\.logger\.Info("模板处理开始", map\[string\]interface{}{/t.logger.Info("模板处理开始，模板ID: %s", t.ID)/' template.go
sed -i 's/t\.logger\.Info("模板处理完成", map\[string\]interface{}{/t.logger.Info("模板处理完成，模板ID: %s", t.ID)/' template.go
sed -i 's/t\.logger\.Info("占位符替换开始", map\[string\]interface{}{/t.logger.Info("占位符替换开始，占位符数量: %d", len(t.Placeholders))/' template.go
sed -i 's/t\.logger\.Info("占位符替换完成", map\[string\]interface{}{/t.logger.Info("占位符替换完成，占位符数量: %d", len(t.Placeholders))/' template.go

# 修复 template_enhanced.go
sed -i 's/et\.logger\.Info("增强模板处理开始", map\[string\]interface{}{/et.logger.Info("增强模板处理开始，模板ID: %s", et.ID)/' template_enhanced.go
sed -i 's/et\.logger\.Info("增强模板处理完成", map\[string\]interface{}{/et.logger.Info("增强模板处理完成，模板ID: %s", et.ID)/' template_enhanced.go
sed -i 's/et\.logger\.Info("变量替换开始", map\[string\]interface{}{/et.logger.Info("变量替换开始，变量数量: %d", len(et.Variables))/' template_enhanced.go
sed -i 's/et\.logger\.Info("变量替换完成", map\[string\]interface{}{/et.logger.Info("变量替换完成，变量数量: %d", len(et.Variables))/' template_enhanced.go
sed -i 's/et\.logger\.Info("条件处理开始", map\[string\]interface{}{/et.logger.Info("条件处理开始，条件数量: %d", len(et.Conditions))/' template_enhanced.go
sed -i 's/et\.logger\.Info("条件处理完成", map\[string\]interface{}{/et.logger.Info("条件处理完成，条件数量: %d", len(et.Conditions))/' template_enhanced.go
sed -i 's/et\.logger\.Info("循环处理开始", map\[string\]interface{}{/et.logger.Info("循环处理开始，循环数量: %d", len(et.Loops))/' template_enhanced.go
sed -i 's/et\.logger\.Info("循环处理完成", map\[string\]interface{}{/et.logger.Info("循环处理完成，循环数量: %d", len(et.Loops))/' template_enhanced.go
sed -i 's/et\.logger\.Info("增强模板验证开始", map\[string\]interface{}{/et.logger.Info("增强模板验证开始，模板ID: %s", et.ID)/' template_enhanced.go
sed -i 's/et\.logger\.Info("增强模板验证完成", map\[string\]interface{}{/et.logger.Info("增强模板验证完成，模板ID: %s", et.ID)/' template_enhanced.go
sed -i 's/et\.logger\.Info("增强模板处理失败", map\[string\]interface{}{/et.logger.Info("增强模板处理失败，模板ID: %s, 错误: %v", et.ID, err)/' template_enhanced.go
sed -i 's/et\.logger\.Info("增强模板处理警告", map\[string\]interface{}{/et.logger.Info("增强模板处理警告，模板ID: %s, 警告: %s", et.ID, warning)/' template_enhanced.go

echo "Logger调用格式修复完成！"
