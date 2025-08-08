package tests

import (
	"context"
	"fmt"
	"runtime"
	"sync"
	"testing"
	"time"

	"github.com/tanqiangyes/go-word/pkg/utils"
)

// TestMemoryPool tests memory pool functionality
func TestMemoryPool(t *testing.T) {
	// Create memory pool for strings
	pool := utils.NewMemoryPool(func() interface{} {
		return ""
	})

	// Test basic pool operations
	obj1 := pool.Get().(string)
	if obj1 != "" {
		t.Errorf("Expected empty string, got: %s", obj1)
	}

	pool.Put("test string")
	obj2 := pool.Get().(string)
	if obj2 != "test string" {
		t.Errorf("Expected 'test string', got: %s", obj2)
	}

	// Test metrics
	metrics := pool.GetMetrics()
	if metrics.Allocations < 1 {
		t.Error("Expected allocations to be at least 1")
	}
}

// TestCache tests cache functionality
func TestCache(t *testing.T) {
	cache := utils.NewCache()

	// Test basic cache operations
	cache.Set("key1", "value1", time.Second)
	
	value, exists := cache.Get("key1")
	if !exists {
		t.Error("Expected key1 to exist")
	}
	if value != "value1" {
		t.Errorf("Expected 'value1', got: %v", value)
	}

	// Test cache miss
	_, exists = cache.Get("nonexistent")
	if exists {
		t.Error("Expected key to not exist")
	}

	// Test expiration
	cache.Set("expire", "value", time.Millisecond)
	time.Sleep(10 * time.Millisecond)
	_, exists = cache.Get("expire")
	if exists {
		t.Error("Expected expired key to not exist")
	}

	// Test metrics
	metrics := cache.GetMetrics()
	if metrics.Hits < 1 {
		t.Error("Expected at least 1 hit")
	}
}

// TestConcurrencyManager tests concurrency manager functionality
func TestConcurrencyManager(t *testing.T) {
	manager := utils.NewConcurrencyManager(2, 10*time.Millisecond)
	ctx := context.Background()

	var wg sync.WaitGroup
	var completed int
	var mu sync.Mutex

	// Test concurrent execution
	for i := 0; i < 5; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()
			
			err := manager.ExecuteWithLimit(ctx, func() error {
				mu.Lock()
				completed++
				mu.Unlock()
				time.Sleep(10 * time.Millisecond)
				return nil
			})
			
			if err != nil {
				t.Errorf("Task %d failed: %v", id, err)
			}
		}(i)
	}

	wg.Wait()

	if completed != 5 {
		t.Errorf("Expected 5 completed tasks, got: %d", completed)
	}

	// Test metrics
	metrics := manager.GetMetrics()
	if metrics.CompletedTasks != 5 {
		t.Errorf("Expected 5 completed tasks, got: %d", metrics.CompletedTasks)
	}
}

// TestResourceManager tests resource manager functionality
func TestResourceManager(t *testing.T) {
	manager := utils.NewResourceManager(1024*1024, 50.0) // 1MB, 50% CPU

	// Monitor resources
	manager.MonitorResources()

	// Test metrics
	metrics := manager.GetMetrics()
	if metrics.MemoryUsage == 0 {
		t.Error("Expected memory usage to be greater than 0")
	}
	if metrics.GoroutineCount == 0 {
		t.Error("Expected goroutine count to be greater than 0")
	}

	// Test alerts
	alerts := manager.GetAlerts()
	// Alerts might be empty depending on current system state
	_ = alerts // Just check that it doesn't panic
}

// TestPerformanceOptimizerAdvanced tests advanced performance optimizer
func TestPerformanceOptimizerAdvanced(t *testing.T) {
	optimizer := utils.NewPerformanceOptimizerAdvanced()

	// Test document processing optimization
	documents := []string{"doc1", "doc2", "doc3"}
	
	ctx := context.Background()
	err := optimizer.OptimizeDocumentProcessing(ctx, documents, func(doc string) error {
		// Simulate document processing
		time.Sleep(10 * time.Millisecond)
		return nil
	})

	if err != nil {
		t.Errorf("Document processing failed: %v", err)
	}

	// Test optimization report
	report := optimizer.GetOptimizationReport()
	if report == "" {
		t.Error("Expected non-empty optimization report")
	}
}

// TestMemoryOptimization tests memory optimization utilities
func TestMemoryOptimization(t *testing.T) {
	// Test memory optimization
	utils.OptimizeMemoryUsage()

	// Test concurrency optimization
	utils.OptimizeConcurrency(runtime.NumCPU())

	// Test performance tips
	tips := utils.GetAdvancedPerformanceTips()
	if len(tips) == 0 {
		t.Error("Expected non-empty performance tips")
	}
}

// TestConcurrentDocumentProcessing tests concurrent document processing
func TestConcurrentDocumentProcessing(t *testing.T) {
	optimizer := utils.NewPerformanceOptimizerAdvanced()
	
	// Create test documents
	documents := make([]string, 10)
	for i := range documents {
		documents[i] = fmt.Sprintf("test_doc_%d.docx", i)
	}

	// Process documents concurrently
	ctx := context.Background()
	start := time.Now()
	
	err := optimizer.OptimizeDocumentProcessing(ctx, documents, func(doc string) error {
		// Simulate document processing
		time.Sleep(50 * time.Millisecond)
		return nil
	})

	duration := time.Since(start)

	if err != nil {
		t.Errorf("Concurrent processing failed: %v", err)
	}

	// Verify processing time is reasonable (should be much less than sequential)
	expectedSequentialTime := 10 * 50 * time.Millisecond
	if duration > expectedSequentialTime {
		t.Errorf("Concurrent processing took too long: %v (expected less than %v)", 
			duration, expectedSequentialTime)
	}
}

// TestMemoryPoolStress tests memory pool under stress
func TestMemoryPoolStress(t *testing.T) {
	pool := utils.NewMemoryPool(func() interface{} {
		return make([]byte, 1024)
	})

	var wg sync.WaitGroup
	iterations := 1000
	concurrency := 10

	// Stress test the pool
	for i := 0; i < concurrency; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for j := 0; j < iterations; j++ {
				obj := pool.Get().([]byte)
				// Use the object
				for k := range obj {
					obj[k] = byte(k % 256)
				}
				pool.Put(obj)
			}
		}()
	}

	wg.Wait()

	// Check metrics
	metrics := pool.GetMetrics()
	if metrics.Allocations < int64(iterations*concurrency) {
		t.Errorf("Expected at least %d allocations, got: %d", 
			iterations*concurrency, metrics.Allocations)
	}
}

// TestCacheStress tests cache under stress
func TestCacheStress(t *testing.T) {
	cache := utils.NewCache()

	var wg sync.WaitGroup
	iterations := 1000
	concurrency := 10

	// Stress test the cache
	for i := 0; i < concurrency; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()
			for j := 0; j < iterations; j++ {
				key := fmt.Sprintf("key_%d_%d", id, j)
				value := fmt.Sprintf("value_%d_%d", id, j)
				
				cache.Set(key, value, time.Second)
				
				// Retrieve the value
				if retrieved, exists := cache.Get(key); !exists || retrieved != value {
					t.Errorf("Cache miss for key: %s", key)
				}
			}
		}(i)
	}

	wg.Wait()

	// Check metrics
	metrics := cache.GetMetrics()
	if metrics.Hits < int64(iterations*concurrency) {
		t.Errorf("Expected at least %d hits, got: %d", 
			iterations*concurrency, metrics.Hits)
	}
}

// TestResourceMonitoring tests resource monitoring over time
func TestResourceMonitoring(t *testing.T) {
	manager := utils.NewResourceManager(1024*1024, 50.0)

	// Monitor resources multiple times
	for i := 0; i < 5; i++ {
		manager.MonitorResources()
		time.Sleep(10 * time.Millisecond)
	}

	// Check that metrics are being collected
	metrics := manager.GetMetrics()
	if metrics.MemoryUsage == 0 {
		t.Error("Expected memory usage to be tracked")
	}
	if metrics.GoroutineCount == 0 {
		t.Error("Expected goroutine count to be tracked")
	}
}

// TestPerformanceReportGeneration tests performance report generation
func TestPerformanceReportGeneration(t *testing.T) {
	optimizer := utils.NewPerformanceOptimizerAdvanced()

	// Generate some activity
	for i := 0; i < 5; i++ {
		optimizer.OptimizeDocumentProcessing(context.Background(), 
			[]string{fmt.Sprintf("doc_%d.docx", i)}, 
			func(doc string) error {
				time.Sleep(10 * time.Millisecond)
				return nil
			})
	}

	// Generate report
	report := optimizer.GetOptimizationReport()
	
	// Verify report contains expected sections
	expectedSections := []string{
		"资源使用情况:",
		"缓存性能:",
		"并发性能:",
		"对象池性能:",
	}

	for _, section := range expectedSections {
		if !containsPerformanceReport(report, section) {
			t.Errorf("Report missing section: %s", section)
		}
	}
}

// Helper function to check if string contains substring
func containsPerformanceReport(s, substr string) bool {
	return len(s) >= len(substr) && (s == substr || 
		(len(s) > len(substr) && (s[:len(substr)] == substr || 
		s[len(s)-len(substr):] == substr || 
		containsSubstringPerformanceReport(s, substr))))
}

func containsSubstringPerformanceReport(s, substr string) bool {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
} 