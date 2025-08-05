// Package utils provides utility functions for the go-word library.
// This file contains performance monitoring and optimization utilities.

package utils

import (
	"fmt"
	"runtime"
	"time"
)

// PerformanceMetrics tracks performance metrics for document operations
type PerformanceMetrics struct {
	Operation     string
	StartTime     time.Time
	EndTime       time.Time
	Duration      time.Duration
	MemoryBefore  uint64
	MemoryAfter   uint64
	MemoryDelta   int64
	GCCountBefore uint32
	GCCountAfter  uint32
	GCCountDelta  uint32
}

// PerformanceMonitor provides performance monitoring capabilities
type PerformanceMonitor struct {
	metrics []PerformanceMetrics
	enabled bool
}

// NewPerformanceMonitor creates a new performance monitor
func NewPerformanceMonitor(enabled bool) *PerformanceMonitor {
	return &PerformanceMonitor{
		metrics: make([]PerformanceMetrics, 0),
		enabled: enabled,
	}
}

// StartOperation starts monitoring an operation
func (pm *PerformanceMonitor) StartOperation(operation string) *PerformanceMetrics {
	if !pm.enabled {
		return nil
	}

	var m runtime.MemStats
	runtime.ReadMemStats(&m)

	metrics := &PerformanceMetrics{
		Operation:     operation,
		StartTime:     time.Now(),
		MemoryBefore:  m.Alloc,
		GCCountBefore: m.NumGC,
	}

	return metrics
}

// EndOperation ends monitoring an operation
func (pm *PerformanceMonitor) EndOperation(metrics *PerformanceMetrics) {
	if !pm.enabled || metrics == nil {
		return
	}

	var m runtime.MemStats
	runtime.ReadMemStats(&m)

	metrics.EndTime = time.Now()
	metrics.Duration = metrics.EndTime.Sub(metrics.StartTime)
	metrics.MemoryAfter = m.Alloc
	metrics.MemoryDelta = int64(metrics.MemoryAfter) - int64(metrics.MemoryBefore)
	metrics.GCCountAfter = m.NumGC
	metrics.GCCountDelta = metrics.GCCountAfter - metrics.GCCountBefore

	pm.metrics = append(pm.metrics, *metrics)
}

// GetMetrics returns all collected performance metrics
func (pm *PerformanceMonitor) GetMetrics() []PerformanceMetrics {
	return pm.metrics
}

// GetLastMetric returns the most recent performance metric
func (pm *PerformanceMonitor) GetLastMetric() *PerformanceMetrics {
	if len(pm.metrics) == 0 {
		return nil
	}
	return &pm.metrics[len(pm.metrics)-1]
}

// PrintSummary prints a summary of all performance metrics
func (pm *PerformanceMonitor) PrintSummary() {
	if !pm.enabled || len(pm.metrics) == 0 {
		return
	}

	fmt.Println("=== 性能监控摘要 ===")
	
	var totalDuration time.Duration
	var totalMemoryDelta int64
	var totalGCCountDelta uint32

	for _, metric := range pm.metrics {
		fmt.Printf("操作: %s\n", metric.Operation)
		fmt.Printf("  耗时: %v\n", metric.Duration)
		fmt.Printf("  内存变化: %+d bytes (%+.2f MB)\n", 
			metric.MemoryDelta, float64(metric.MemoryDelta)/1024/1024)
		fmt.Printf("  GC次数变化: %+d\n", metric.GCCountDelta)
		fmt.Println()

		totalDuration += metric.Duration
		totalMemoryDelta += metric.MemoryDelta
		totalGCCountDelta += metric.GCCountDelta
	}

	fmt.Printf("总计:\n")
	fmt.Printf("  总耗时: %v\n", totalDuration)
	fmt.Printf("  总内存变化: %+d bytes (%+.2f MB)\n", 
		totalMemoryDelta, float64(totalMemoryDelta)/1024/1024)
	fmt.Printf("  总GC次数变化: %+d\n", totalGCCountDelta)
}

// PerformanceOptimizer provides optimization suggestions
type PerformanceOptimizer struct {
	monitor *PerformanceMonitor
}

// NewPerformanceOptimizer creates a new performance optimizer
func NewPerformanceOptimizer(monitor *PerformanceMonitor) *PerformanceOptimizer {
	return &PerformanceOptimizer{
		monitor: monitor,
	}
}

// AnalyzePerformance analyzes performance metrics and provides suggestions
func (po *PerformanceOptimizer) AnalyzePerformance() []string {
	if po.monitor == nil || len(po.monitor.metrics) == 0 {
		return []string{"没有性能数据可供分析"}
	}

	var suggestions []string

	// Analyze memory usage
	var totalMemoryDelta int64
	for _, metric := range po.monitor.metrics {
		totalMemoryDelta += metric.MemoryDelta
	}

	if totalMemoryDelta > 100*1024*1024 { // 100MB
		suggestions = append(suggestions, 
			"⚠️  内存使用较高，建议在处理大文档时及时关闭文档释放资源")
	}

	// Analyze GC frequency
	var totalGCCount uint32
	for _, metric := range po.monitor.metrics {
		totalGCCount += metric.GCCountDelta
	}

	if totalGCCount > 10 {
		suggestions = append(suggestions, 
			"⚠️  GC频率较高，建议优化内存分配模式")
	}

	// Analyze operation duration
	for _, metric := range po.monitor.metrics {
		if metric.Duration > 5*time.Second {
			suggestions = append(suggestions, 
				fmt.Sprintf("⚠️  操作 '%s' 耗时较长 (%v)，建议检查文档大小和复杂度", 
					metric.Operation, metric.Duration))
		}
	}

	if len(suggestions) == 0 {
		suggestions = append(suggestions, "✅ 性能表现良好，无需优化")
	}

	return suggestions
}

// MemoryProfiler provides memory profiling capabilities
type MemoryProfiler struct {
	snapshots []MemorySnapshot
}

// MemorySnapshot represents a memory usage snapshot
type MemorySnapshot struct {
	Timestamp time.Time
	Operation string
	Alloc     uint64
	TotalAlloc uint64
	Sys       uint64
	NumGC     uint32
}

// NewMemoryProfiler creates a new memory profiler
func NewMemoryProfiler() *MemoryProfiler {
	return &MemoryProfiler{
		snapshots: make([]MemorySnapshot, 0),
	}
}

// TakeSnapshot takes a memory usage snapshot
func (mp *MemoryProfiler) TakeSnapshot(operation string) {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)

	snapshot := MemorySnapshot{
		Timestamp:  time.Now(),
		Operation:  operation,
		Alloc:      m.Alloc,
		TotalAlloc: m.TotalAlloc,
		Sys:        m.Sys,
		NumGC:      m.NumGC,
	}

	mp.snapshots = append(mp.snapshots, snapshot)
}

// GetMemoryReport returns a detailed memory usage report
func (mp *MemoryProfiler) GetMemoryReport() string {
	if len(mp.snapshots) == 0 {
		return "没有内存快照数据"
	}

	var report string
	report += "=== 内存使用报告 ===\n"

	for i, snapshot := range mp.snapshots {
		report += fmt.Sprintf("快照 %d (%s):\n", i+1, snapshot.Operation)
		report += fmt.Sprintf("  当前分配: %.2f MB\n", float64(snapshot.Alloc)/1024/1024)
		report += fmt.Sprintf("  总分配: %.2f MB\n", float64(snapshot.TotalAlloc)/1024/1024)
		report += fmt.Sprintf("  系统内存: %.2f MB\n", float64(snapshot.Sys)/1024/1024)
		report += fmt.Sprintf("  GC次数: %d\n", snapshot.NumGC)
		report += "\n"
	}

	// Calculate memory growth
	if len(mp.snapshots) > 1 {
		first := mp.snapshots[0]
		last := mp.snapshots[len(mp.snapshots)-1]
		
		memoryGrowth := int64(last.Alloc) - int64(first.Alloc)
		report += fmt.Sprintf("内存增长: %+d bytes (%+.2f MB)\n", 
			memoryGrowth, float64(memoryGrowth)/1024/1024)
	}

	return report
}

// Performance utilities

// MeasureOperation measures the performance of an operation
func MeasureOperation(operation string, fn func() error) error {
	monitor := NewPerformanceMonitor(true)
	metrics := monitor.StartOperation(operation)
	
	err := fn()
	
	monitor.EndOperation(metrics)
	
	if err != nil {
		return err
	}
	
	// Print performance summary
	monitor.PrintSummary()
	
	return nil
}

// MeasureOperationWithMemory measures operation performance with memory profiling
func MeasureOperationWithMemory(operation string, fn func() error) error {
	profiler := NewMemoryProfiler()
	profiler.TakeSnapshot("开始")
	
	err := fn()
	
	profiler.TakeSnapshot("结束")
	
	if err != nil {
		return err
	}
	
	// Print memory report
	fmt.Println(profiler.GetMemoryReport())
	
	return nil
}

// GetSystemInfo returns system information for performance analysis
func GetSystemInfo() map[string]interface{} {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)

	return map[string]interface{}{
		"goroutines":    runtime.NumGoroutine(),
		"cpu_count":     runtime.NumCPU(),
		"memory_alloc":  m.Alloc,
		"memory_sys":    m.Sys,
		"gc_count":      m.NumGC,
		"go_version":    runtime.Version(),
	}
}

// PrintSystemInfo prints current system information
func PrintSystemInfo() {
	info := GetSystemInfo()
	
	fmt.Println("=== 系统信息 ===")
	fmt.Printf("Go版本: %s\n", info["go_version"])
	fmt.Printf("CPU核心数: %d\n", info["cpu_count"])
	fmt.Printf("当前Goroutine数: %d\n", info["goroutines"])
	fmt.Printf("内存分配: %.2f MB\n", float64(info["memory_alloc"].(uint64))/1024/1024)
	fmt.Printf("系统内存: %.2f MB\n", float64(info["memory_sys"].(uint64))/1024/1024)
	fmt.Printf("GC次数: %d\n", info["gc_count"])
}

// Performance tips and best practices

// GetPerformanceTips returns performance optimization tips
func GetPerformanceTips() []string {
	return []string{
		"💡 处理大文档时，使用 defer doc.Close() 确保资源及时释放",
		"💡 批量处理文档时，处理完一个文档后立即关闭",
		"💡 避免在循环中重复打开同一个文档",
		"💡 对于只读操作，考虑使用流式处理减少内存占用",
		"💡 定期调用 runtime.GC() 手动触发垃圾回收",
		"💡 使用 sync.Pool 复用对象减少内存分配",
		"💡 考虑使用 goroutine 并发处理多个文档",
		"💡 监控内存使用，避免内存泄漏",
	}
}

// PrintPerformanceTips prints performance optimization tips
func PrintPerformanceTips() {
	fmt.Println("=== 性能优化建议 ===")
	tips := GetPerformanceTips()
	for i, tip := range tips {
		fmt.Printf("%d. %s\n", i+1, tip)
	}
} 