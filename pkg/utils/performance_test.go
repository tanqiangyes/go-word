package utils

import (
	"testing"
	"time"
)

func TestNewPerformanceMonitor(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	if monitor == nil {
		t.Error("Expected non-nil performance monitor")
	}

	if monitor.metrics == nil {
		t.Error("Expected metrics slice to be initialized")
	}

	if !monitor.enabled {
		t.Error("Expected monitor to be enabled")
	}
}

func TestPerformanceMonitor_StartOperation(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	
	// Start an operation
	metrics := monitor.StartOperation("test_operation")
	if metrics == nil {
		t.Error("Expected non-nil metrics")
	}
	
	if metrics.Operation != "test_operation" {
		t.Errorf("Expected operation name 'test_operation', got '%s'", metrics.Operation)
	}
	
	if metrics.StartTime.IsZero() {
		t.Error("Expected start time to be set")
	}
	
	if metrics.MemoryBefore == 0 {
		t.Error("Expected memory before to be set")
	}
}

func TestPerformanceMonitor_EndOperation(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	
	// Start and end an operation
	metrics := monitor.StartOperation("test_operation")
	time.Sleep(10 * time.Millisecond) // Small delay to ensure measurable time
	monitor.EndOperation(metrics)
	
	// Check that the operation was recorded
	allMetrics := monitor.GetMetrics()
	if len(allMetrics) != 1 {
		t.Errorf("Expected 1 metric, got %d", len(allMetrics))
	}
	
	metric := allMetrics[0]
	if metric.Operation != "test_operation" {
		t.Errorf("Expected metric operation 'test_operation', got '%s'", metric.Operation)
	}
	
	if metric.Duration <= 0 {
		t.Error("Expected positive duration")
	}
	
	if metric.StartTime.IsZero() {
		t.Error("Expected start time to be set")
	}
	
	if metric.EndTime.IsZero() {
		t.Error("Expected end time to be set")
	}
	
	if metric.MemoryDelta == 0 {
		t.Error("Expected memory delta to be calculated")
	}
}

func TestPerformanceMonitor_EndOperationWithoutStart(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	
	// Try to end an operation that wasn't started
	monitor.EndOperation(nil)
	
	// Should not cause panic and should not add metrics
	metrics := monitor.GetMetrics()
	if len(metrics) != 0 {
		t.Errorf("Expected 0 metrics, got %d", len(metrics))
	}
}

func TestPerformanceMonitor_GetMetrics(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	
	// Add multiple operations
	metrics1 := monitor.StartOperation("op1")
	time.Sleep(5 * time.Millisecond)
	monitor.EndOperation(metrics1)
	
	metrics2 := monitor.StartOperation("op2")
	time.Sleep(5 * time.Millisecond)
	monitor.EndOperation(metrics2)
	
	allMetrics := monitor.GetMetrics()
	if len(allMetrics) != 2 {
		t.Errorf("Expected 2 metrics, got %d", len(allMetrics))
	}
	
	// Check that metrics are in chronological order
	if allMetrics[0].StartTime.After(allMetrics[1].StartTime) {
		t.Error("Expected metrics to be in chronological order")
	}
}

func TestPerformanceMonitor_GetLastMetric(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	
	// No metrics yet
	lastMetric := monitor.GetLastMetric()
	if lastMetric != nil {
		t.Error("Expected nil for last metric when no metrics exist")
	}
	
	// Add a metric
	metrics := monitor.StartOperation("test_operation")
	time.Sleep(5 * time.Millisecond)
	monitor.EndOperation(metrics)
	
	lastMetric = monitor.GetLastMetric()
	if lastMetric == nil {
		t.Error("Expected non-nil last metric")
	}
	
	if lastMetric.Operation != "test_operation" {
		t.Errorf("Expected last metric operation 'test_operation', got '%s'", lastMetric.Operation)
	}
}

func TestPerformanceMonitor_PrintSummary(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	
	// Add some operations
	metrics1 := monitor.StartOperation("fast_operation")
	time.Sleep(5 * time.Millisecond)
	monitor.EndOperation(metrics1)
	
	metrics2 := monitor.StartOperation("slow_operation")
	time.Sleep(15 * time.Millisecond)
	monitor.EndOperation(metrics2)
	
	// This should not panic and should print something
	monitor.PrintSummary()
}

func TestNewPerformanceOptimizer(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	optimizer := NewPerformanceOptimizer(monitor)
	if optimizer == nil {
		t.Error("Expected non-nil performance optimizer")
	}
}

func TestPerformanceOptimizer_AnalyzePerformance(t *testing.T) {
	monitor := NewPerformanceMonitor(true)
	optimizer := NewPerformanceOptimizer(monitor)
	
	// Add some test metrics
	metrics1 := monitor.StartOperation("fast_op")
	time.Sleep(5 * time.Millisecond)
	monitor.EndOperation(metrics1)
	
	metrics2 := monitor.StartOperation("slow_op")
	time.Sleep(50 * time.Millisecond)
	monitor.EndOperation(metrics2)
	
	// Analyze performance
	suggestions := optimizer.AnalyzePerformance()
	
	if len(suggestions) == 0 {
		t.Error("Expected performance suggestions")
	}
	
	// Check that suggestions contain useful information
	foundUseful := false
	for _, suggestion := range suggestions {
		if len(suggestion) > 10 {
			foundUseful = true
			break
		}
	}
	
	if !foundUseful {
		t.Error("Expected useful performance suggestions")
	}
}

func TestNewMemoryProfiler(t *testing.T) {
	profiler := NewMemoryProfiler()
	if profiler == nil {
		t.Error("Expected non-nil memory profiler")
	}
}

func TestMemoryProfiler_TakeSnapshot(t *testing.T) {
	profiler := NewMemoryProfiler()
	
	profiler.TakeSnapshot("test_snapshot")
	
	// Check that snapshot was taken
	snapshots := profiler.snapshots
	if len(snapshots) != 1 {
		t.Errorf("Expected 1 snapshot, got %d", len(snapshots))
	}
	
	snapshot := snapshots[0]
	if snapshot.Operation != "test_snapshot" {
		t.Errorf("Expected snapshot operation 'test_snapshot', got '%s'", snapshot.Operation)
	}
	
	if snapshot.Timestamp.IsZero() {
		t.Error("Expected timestamp to be set")
	}
	
	if snapshot.Alloc <= 0 {
		t.Error("Expected positive memory allocation")
	}
	
	if snapshot.NumGC < 0 {
		t.Error("Expected non-negative GC count")
	}
}

func TestMemoryProfiler_GetMemoryReport(t *testing.T) {
	profiler := NewMemoryProfiler()
	
	// Take multiple snapshots
	profiler.TakeSnapshot("snapshot1")
	time.Sleep(10 * time.Millisecond)
	profiler.TakeSnapshot("snapshot2")
	time.Sleep(10 * time.Millisecond)
	profiler.TakeSnapshot("snapshot3")
	
	report := profiler.GetMemoryReport()
	if report == "" {
		t.Error("Expected non-empty memory report")
	}
	
	// Check that report contains expected information
	if !contains(report, "snapshot1") {
		t.Error("Expected report to contain snapshot1")
	}
	
	if !contains(report, "snapshot2") {
		t.Error("Expected report to contain snapshot2")
	}
	
	if !contains(report, "snapshot3") {
		t.Error("Expected report to contain snapshot3")
	}
}

func TestMeasureOperation(t *testing.T) {
	// Test measuring a simple operation
	err := MeasureOperation("test_measure", func() error {
		time.Sleep(10 * time.Millisecond)
		return nil
	})
	
	if err != nil {
		t.Errorf("Expected no error, got %v", err)
	}
}

func TestMeasureOperationWithMemory(t *testing.T) {
	// Test measuring operation with memory tracking
	err := MeasureOperationWithMemory("test_memory_measure", func() error {
		time.Sleep(10 * time.Millisecond)
		return nil
	})
	
	if err != nil {
		t.Errorf("Expected no error, got %v", err)
	}
}

func TestGetSystemInfo(t *testing.T) {
	info := GetSystemInfo()
	if info == nil {
		t.Error("Expected non-nil system info")
	}
	
	// Check that system info contains expected keys
	expectedKeys := []string{"goroutines", "cpu_count", "memory_alloc", "memory_sys", "gc_count", "go_version"}
	for _, key := range expectedKeys {
		if _, exists := info[key]; !exists {
			t.Errorf("Expected system info to contain key '%s'", key)
		}
	}
}

func TestPrintSystemInfo(t *testing.T) {
	// This should not panic
	PrintSystemInfo()
}

func TestGetPerformanceTips(t *testing.T) {
	tips := GetPerformanceTips()
	if len(tips) == 0 {
		t.Error("Expected performance tips")
	}
	
	// Check that tips contain useful information
	foundMemory := false
	foundGoroutine := false
	for _, tip := range tips {
		if contains(tip, "内存") {
			foundMemory = true
		}
		if contains(tip, "goroutine") {
			foundGoroutine = true
		}
	}
	
	if !foundMemory {
		t.Error("Expected tips to mention memory optimization")
	}
	
	if !foundGoroutine {
		t.Error("Expected tips to mention goroutine usage")
	}
}

func TestPrintPerformanceTips(t *testing.T) {
	// This should not panic
	PrintPerformanceTips()
}

// Helper function to check if a string contains a substring
func contains(s, substr string) bool {
	return len(s) >= len(substr) && (s == substr || 
		(len(s) > len(substr) && (s[:len(substr)] == substr || 
		s[len(s)-len(substr):] == substr || 
		containsSubstring(s, substr))))
}

func containsSubstring(s, substr string) bool {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
} 