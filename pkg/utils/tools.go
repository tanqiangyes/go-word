package utils

import (
	"crypto/rand"
	"encoding/hex"
	"sort"
	"time"
)

// GenerateID 生成唯一ID
func GenerateID() string {
	bytes := make([]byte, 16)
	rand.Read(bytes)
	return hex.EncodeToString(bytes)
}

// GetCurrentTimestamp 获取当前时间戳
func GetCurrentTimestamp() int64 {
	return time.Now().Unix()
}

// SortByTimestamp 按时间戳排序（最新的在前）
func SortByTimestamp[T any](items []T, getTimestamp func(T) time.Time) {
	sort.Slice(items, func(i, j int) bool {
		return getTimestamp(items[i]).After(getTimestamp(items[j]))
	})
}

// SortByPriorityAndTimestamp 按优先级和时间戳排序
func SortByPriorityAndTimestamp[T any](items []T, getPriorityAndTimestamp func(T) (string, time.Time)) {
	sort.Slice(items, func(i, j int) bool {
		priorityI, timestampI := getPriorityAndTimestamp(items[i])
		priorityJ, timestampJ := getPriorityAndTimestamp(items[j])
		
		// 优先级比较
		if priorityI != priorityJ {
			return getPriorityWeight(priorityI) > getPriorityWeight(priorityJ)
		}
		
		// 时间戳比较（最新的在前）
		return timestampI.After(timestampJ)
	})
}

// getPriorityWeight 获取优先级权重
func getPriorityWeight(priority string) int {
	switch priority {
	case "critical":
		return 4
	case "high":
		return 3
	case "medium":
		return 2
	case "low":
		return 1
	default:
		return 0
	}
}
