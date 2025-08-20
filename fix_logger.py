#!/usr/bin/env python3
import os
import re
import glob

def fix_logger_calls(file_path):
    """修复文件中的Logger调用格式"""
    print(f"修复文件: {file_path}")
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 修复 logger.Info 调用
    # 匹配模式: logger.Info("消息", map[string]interface{}{...})
    pattern = r'(\w+\.logger\.Info\("([^"]+)",\s*map\[string\]interface\{\}\s*\{([^}]+)\})'
    
    def replace_logger_call(match):
        full_call = match.group(1)
        message = match.group(2)
        fields = match.group(3)
        
        # 解析字段
        field_pairs = []
        for line in fields.split('\n'):
            line = line.strip()
            if ':' in line and not line.startswith('//'):
                parts = line.split(':')
                if len(parts) >= 2:
                    key = parts[0].strip().strip('"')
                    value = parts[1].strip().rstrip(',').strip()
                    field_pairs.append((key, value))
        
        if not field_pairs:
            return full_call
        
        # 构建新的格式化字符串
        format_args = []
        new_message = message
        
        for key, value in field_pairs:
            if value in ['original.ID', 'new.ID', 'mergedStyle.ID', 'styleID', 'conflict.ID', 'operation.ID', 'session.ID', 'userID', 'sessionID', 'controlID', 'fileID', 'linkID', 'containerID', 'tab.ID', 'group.ID', 'control.ID', 'comment.ID', 'image.ID', 'table.ID', 'paragraph.ID', 'protectionType', 'validationType', 'format', 'subject', 'keywords', 'author', 'title', 'filePath', 'outputPath']:
                if 'ID' in value:
                    new_message += f", {key}: %s"
                    format_args.append(value)
                elif 'Type' in value:
                    new_message += f", {key}: %s"
                    format_args.append(value)
                elif 'Path' in value:
                    new_message += f", {key}: %s"
                    format_args.append(value)
                elif value in ['title', 'author', 'subject', 'keywords', 'format']:
                    new_message += f", {key}: %s"
                    format_args.append(value)
                else:
                    new_message += f", {key}: %s"
                    format_args.append(value)
            elif 'len(' in value:
                new_message += f", {key}: %d"
                format_args.append(value)
            elif value.isdigit() or value.replace('.', '').isdigit():
                new_message += f", {key}: %d"
                format_args.append(value)
            elif value == 'true' or value == 'false':
                new_message += f", {key}: %t"
                format_args.append(value)
            else:
                new_message += f", {key}: %s"
                format_args.append(value)
        
        # 构建新的调用
        if format_args:
            args_str = ', '.join(format_args)
            return f'{match.group(0).split(".")[0]}.logger.Info("{new_message}", {args_str})'
        else:
            return f'{match.group(0).split(".")[0]}.logger.Info("{new_message}")'
    
    # 应用替换
    new_content = re.sub(pattern, replace_logger_call, content, flags=re.MULTILINE | re.DOTALL)
    
    # 修复 logger.Debug 调用
    pattern = r'(\w+\.logger\.Debug\("([^"]+)",\s*map\[string\]interface\{\}\s*\{([^}]+)\})'
    new_content = re.sub(pattern, replace_logger_call, new_content, flags=re.MULTILINE | re.DOTALL)
    
    # 修复 logger.Warning 调用
    pattern = r'(\w+\.logger\.Warning\("([^"]+)",\s*map\[string\]interface\{\}\s*\{([^}]+)\})'
    new_content = re.sub(pattern, replace_logger_call, new_content, flags=re.MULTILINE | re.DOTALL)
    
    # 修复 logger.Error 调用
    pattern = r'(\w+\.logger\.Error\("([^"]+)",\s*map\[string\]interface\{\}\s*\{([^}]+)\})'
    new_content = re.sub(pattern, replace_logger_call, new_content, flags=re.MULTILINE | re.DOTALL)
    
    # 如果内容有变化，写回文件
    if new_content != content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"已修复: {file_path}")
        return True
    else:
        print(f"无需修复: {file_path}")
        return False

def main():
    """主函数"""
    print("开始修复Logger调用格式...")
    
    # 获取所有Go文件
    go_files = glob.glob("*.go")
    
    fixed_count = 0
    for go_file in go_files:
        if fix_logger_calls(go_file):
            fixed_count += 1
    
    print(f"\n修复完成！共修复了 {fixed_count} 个文件。")

if __name__ == "__main__":
    main()
