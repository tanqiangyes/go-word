#!/bin/bash

# 删除所有测试文件中的重复函数
for file in tests/*_test.go; do
    if [ "$file" != "tests/test_utils.go" ]; then
        # 删除contains和containsSubstring函数
        sed -i '/^func contains(/,/^}$/d' "$file"
        sed -i '/^func containsSubstring(/,/^}$/d' "$file"
        
        # 删除重复的TestMergeCells函数（除了advanced_formatting_test.go）
        if [ "$file" != "tests/advanced_formatting_test.go" ]; then
            sed -i '/^func TestMergeCells(/,/^}$/d' "$file"
        fi
    fi
done 