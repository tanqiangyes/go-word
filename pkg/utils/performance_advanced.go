// Package utils provides advanced performance optimization utilities for the go-word library.
// This file contains advanced performance optimization features including memory pools,
// caching, concurrency optimization, and resource management.

package utils

import (
	"context"
	"fmt"
	"runtime"
	"sync"
	"sync/atomic"
	"time"

	"github.com/tanqiangyes/go-word/pkg/types"
)

// MemoryPool provides object pooling for frequently allocated objects
type MemoryPool struct {
	pool    sync.Pool
	metrics *PoolMetrics
}

// PoolMetrics tracks memory pool usage statistics
type PoolMetrics struct {
	Allocations   int64
	Reuses        int64
	Misses        int64
	TotalGetTime  time.Duration
	TotalPutTime  time.Duration
}

// NewMemoryPool creates a new memory pool for the specified object type
func NewMemoryPool(newFunc func() interface{}) *MemoryPool {
	return &MemoryPool{
		pool: sync.Pool{
			New: newFunc,
		},
		metrics: &PoolMetrics{},
	}
}

// Get retrieves an object from the pool
func (mp *MemoryPool) Get() interface{} {
	start := time.Now()
	obj := mp.pool.Get()
	mp.metrics.TotalGetTime += time.Since(start)
	
	if obj == nil {
		atomic.AddInt64(&mp.metrics.Misses, 1)
	} else {
		atomic.AddInt64(&mp.metrics.Reuses, 1)
	}
	atomic.AddInt64(&mp.metrics.Allocations, 1)
	
	return obj
}

// Put returns an object to the pool
func (mp *MemoryPool) Put(obj interface{}) {
	start := time.Now()
	mp.pool.Put(obj)
	mp.metrics.TotalPutTime += time.Since(start)
}

// GetMetrics returns pool usage metrics
func (mp *MemoryPool) GetMetrics() PoolMetrics {
	return *mp.metrics
}

// Cache provides a thread-safe cache with TTL support
type Cache struct {
	data    map[string]CacheEntry
	mutex   sync.RWMutex
	metrics *CacheMetrics
}

// CacheEntry represents a cached item
type CacheEntry struct {
	Value      interface{}
	ExpiresAt  time.Time
	AccessTime time.Time
}

// CacheMetrics tracks cache usage statistics
type CacheMetrics struct {
	Hits       int64
	Misses     int64
	Evictions  int64
	Size       int64
}

// NewCache creates a new cache instance
func NewCache() *Cache {
	cache := &Cache{
		data:    make(map[string]CacheEntry),
		metrics: &CacheMetrics{},
	}
	
	// Start cleanup goroutine
	go cache.cleanup()
	
	return cache
}

// Set stores a value in the cache with optional TTL
func (c *Cache) Set(key string, value interface{}, ttl time.Duration) {
	c.mutex.Lock()
	defer c.mutex.Unlock()
	
	expiresAt := time.Now().Add(ttl)
	c.data[key] = CacheEntry{
		Value:      value,
		ExpiresAt:  expiresAt,
		AccessTime: time.Now(),
	}
	
	atomic.StoreInt64(&c.metrics.Size, int64(len(c.data)))
}

// Get retrieves a value from the cache
func (c *Cache) Get(key string) (interface{}, bool) {
	c.mutex.RLock()
	entry, exists := c.data[key]
	c.mutex.RUnlock()
	
	if !exists {
		atomic.AddInt64(&c.metrics.Misses, 1)
		return nil, false
	}
	
	// Check if expired
	if time.Now().After(entry.ExpiresAt) {
		c.mutex.Lock()
		delete(c.data, key)
		c.mutex.Unlock()
		atomic.AddInt64(&c.metrics.Evictions, 1)
		atomic.AddInt64(&c.metrics.Misses, 1)
		return nil, false
	}
	
	// Update access time
	c.mutex.Lock()
	if e, exists := c.data[key]; exists {
		entry.AccessTime = time.Now()
		c.data[key] = e
	}
	c.mutex.Unlock()
	
	atomic.AddInt64(&c.metrics.Hits, 1)
	return entry.Value, true
}

// cleanup removes expired entries from the cache
func (c *Cache) cleanup() {
	ticker := time.NewTicker(time.Minute)
	defer ticker.Stop()
	
	for range ticker.C {
		c.mutex.Lock()
		now := time.Now()
		for key, entry := range c.data {
			if now.After(entry.ExpiresAt) {
				delete(c.data, key)
				atomic.AddInt64(&c.metrics.Evictions, 1)
			}
		}
		atomic.StoreInt64(&c.metrics.Size, int64(len(c.data)))
		c.mutex.Unlock()
	}
}

// GetMetrics returns cache usage metrics
func (c *Cache) GetMetrics() CacheMetrics {
	return *c.metrics
}

// ConcurrencyManager manages concurrent operations with rate limiting and resource control
type ConcurrencyManager struct {
	semaphore chan struct{}
    rateLimit *time.Ticker
	metrics   *ConcurrencyMetrics
}

// ConcurrencyMetrics tracks concurrency usage statistics
type ConcurrencyMetrics struct {
	ActiveWorkers    int64
	TotalWorkers     int64
	CompletedTasks   int64
	FailedTasks      int64
	AverageWaitTime  time.Duration
}

// NewConcurrencyManager creates a new concurrency manager
func NewConcurrencyManager(maxWorkers int, rateLimit time.Duration) *ConcurrencyManager {
    cm := &ConcurrencyManager{
        semaphore: make(chan struct{}, maxWorkers),
        metrics:   &ConcurrencyMetrics{},
    }
    if rateLimit > 0 {
        cm.rateLimit = time.NewTicker(rateLimit)
    } else {
        cm.rateLimit = nil
    }
    return cm
}

// ExecuteWithLimit executes a task with concurrency control
func (cm *ConcurrencyManager) ExecuteWithLimit(ctx context.Context, task func() error) error {
	start := time.Now()
	
	// Wait for available slot
	select {
	case cm.semaphore <- struct{}{}:
		defer func() {
			<-cm.semaphore
			atomic.AddInt64(&cm.metrics.ActiveWorkers, -1)
		}()
	case <-ctx.Done():
		return ctx.Err()
	}
	
	atomic.AddInt64(&cm.metrics.ActiveWorkers, 1)
	atomic.AddInt64(&cm.metrics.TotalWorkers, 1)
	
    // Wait for rate limit if enabled
    if cm.rateLimit != nil {
        <-cm.rateLimit.C
    }
	
	// Execute task
	err := task()
	
	// Update metrics
	waitTime := time.Since(start)
	cm.metrics.AverageWaitTime = (cm.metrics.AverageWaitTime + waitTime) / 2
	
	if err != nil {
		atomic.AddInt64(&cm.metrics.FailedTasks, 1)
	} else {
		atomic.AddInt64(&cm.metrics.CompletedTasks, 1)
	}
	
	return err
}

// GetMetrics returns concurrency usage metrics
func (cm *ConcurrencyManager) GetMetrics() ConcurrencyMetrics {
	return *cm.metrics
}

// ResourceManager manages system resources and provides optimization recommendations
type ResourceManager struct {
	memoryThreshold uint64
	cpuThreshold    float64
	alerts          []ResourceAlert
	metrics         *ResourceMetrics
}

// ResourceAlert represents a resource usage alert
type ResourceAlert struct {
	Type      string
	Message   string
	Severity  string
	Timestamp time.Time
}

// ResourceMetrics tracks resource usage
type ResourceMetrics struct {
	MemoryUsage    uint64
	CPUUsage       float64
	GoroutineCount int
	HeapAlloc      uint64
	HeapSys        uint64
}

// NewResourceManager creates a new resource manager
func NewResourceManager(memoryThreshold uint64, cpuThreshold float64) *ResourceManager {
	return &ResourceManager{
		memoryThreshold: memoryThreshold,
		cpuThreshold:    cpuThreshold,
		alerts:          make([]ResourceAlert, 0),
		metrics:         &ResourceMetrics{},
	}
}

// MonitorResources monitors system resources and generates alerts
func (rm *ResourceManager) MonitorResources() {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	
	rm.metrics.MemoryUsage = m.Alloc
	rm.metrics.HeapAlloc = m.HeapAlloc
	rm.metrics.HeapSys = m.HeapSys
	rm.metrics.GoroutineCount = runtime.NumGoroutine()
	
	// Check memory usage
	if m.Alloc > rm.memoryThreshold {
		alert := ResourceAlert{
			Type:      "memory",
			Message:   fmt.Sprintf("内存使用超过阈值: %.2f MB", float64(m.Alloc)/1024/1024),
			Severity:  "warning",
			Timestamp: time.Now(),
		}
		rm.alerts = append(rm.alerts, alert)
	}
	
	// Check goroutine count
	if runtime.NumGoroutine() > 1000 {
		alert := ResourceAlert{
			Type:      "goroutines",
			Message:   fmt.Sprintf("Goroutine数量过多: %d", runtime.NumGoroutine()),
			Severity:  "warning",
			Timestamp: time.Now(),
		}
		rm.alerts = append(rm.alerts, alert)
	}
}

// GetAlerts returns resource usage alerts
func (rm *ResourceManager) GetAlerts() []ResourceAlert {
	return rm.alerts
}

// GetMetrics returns current resource metrics
func (rm *ResourceManager) GetMetrics() ResourceMetrics {
	return *rm.metrics
}

// PerformanceOptimizerAdvanced provides advanced performance optimization features
type PerformanceOptimizerAdvanced struct {
	resourceManager *ResourceManager
	cache           *Cache
	concurrencyMgr  *ConcurrencyManager
	documentPool    *MemoryPool
	paragraphPool   *MemoryPool
	tablePool       *MemoryPool
}

// NewPerformanceOptimizerAdvanced creates a new advanced performance optimizer
func NewPerformanceOptimizerAdvanced() *PerformanceOptimizerAdvanced {
	return &PerformanceOptimizerAdvanced{
		resourceManager: NewResourceManager(100*1024*1024, 80.0), // 100MB, 80% CPU
		cache:           NewCache(),
		// 禁用速率限制以最大化并发性能
		concurrencyMgr:  NewConcurrencyManager(10, 0),
		documentPool:    NewMemoryPool(func() interface{} { return make(map[string]interface{}) }),
		paragraphPool:   NewMemoryPool(func() interface{} { return &types.Paragraph{} }),
		tablePool:       NewMemoryPool(func() interface{} { return &types.Table{} }),
	}
}

// OptimizeDocumentProcessing optimizes document processing with advanced features
func (poa *PerformanceOptimizerAdvanced) OptimizeDocumentProcessing(ctx context.Context, 
	documents []string, processor func(string) error) error {
    // Monitor resources
    poa.resourceManager.MonitorResources()

    // Process documents with concurrency control in parallel
    var wg sync.WaitGroup
    var firstErr error
    var mu sync.Mutex

    for _, doc := range documents {
        d := doc
        wg.Add(1)
        go func() {
            defer wg.Done()
            if err := poa.concurrencyMgr.ExecuteWithLimit(ctx, func() error {
                return processor(d)
            }); err != nil {
                mu.Lock()
                if firstErr == nil {
                    firstErr = err
                }
                mu.Unlock()
            }
        }()
    }

    wg.Wait()
    return firstErr
}

// GetOptimizationReport returns a comprehensive optimization report
func (poa *PerformanceOptimizerAdvanced) GetOptimizationReport() string {
	report := "=== 高级性能优化报告 ===\n\n"
	
	// Resource metrics
	resourceMetrics := poa.resourceManager.GetMetrics()
	report += "资源使用情况:\n"
	report += fmt.Sprintf("  内存使用: %.2f MB\n", float64(resourceMetrics.MemoryUsage)/1024/1024)
	report += fmt.Sprintf("  堆分配: %.2f MB\n", float64(resourceMetrics.HeapAlloc)/1024/1024)
	report += fmt.Sprintf("  系统内存: %.2f MB\n", float64(resourceMetrics.HeapSys)/1024/1024)
	report += fmt.Sprintf("  Goroutine数: %d\n", resourceMetrics.GoroutineCount)
	report += "\n"
	
	// Cache metrics
	cacheMetrics := poa.cache.GetMetrics()
	hitRate := float64(0)
	if cacheMetrics.Hits+cacheMetrics.Misses > 0 {
		hitRate = float64(cacheMetrics.Hits) / float64(cacheMetrics.Hits+cacheMetrics.Misses) * 100
	}
	report += fmt.Sprintf("缓存性能:\n")
	report += fmt.Sprintf("  命中率: %.2f%%\n", hitRate)
	report += fmt.Sprintf("  命中次数: %d\n", cacheMetrics.Hits)
	report += fmt.Sprintf("  未命中次数: %d\n", cacheMetrics.Misses)
	report += fmt.Sprintf("  缓存大小: %d\n", cacheMetrics.Size)
	report += "\n"
	
	// Concurrency metrics
	concurrencyMetrics := poa.concurrencyMgr.GetMetrics()
	report += fmt.Sprintf("并发性能:\n")
	report += fmt.Sprintf("  活跃工作线程: %d\n", concurrencyMetrics.ActiveWorkers)
	report += fmt.Sprintf("  总工作线程: %d\n", concurrencyMetrics.TotalWorkers)
	report += fmt.Sprintf("  完成任务: %d\n", concurrencyMetrics.CompletedTasks)
	report += fmt.Sprintf("  失败任务: %d\n", concurrencyMetrics.FailedTasks)
	report += fmt.Sprintf("  平均等待时间: %v\n", concurrencyMetrics.AverageWaitTime)
	report += "\n"
	
	// Pool metrics
	documentPoolMetrics := poa.documentPool.GetMetrics()
	report += fmt.Sprintf("对象池性能:\n")
	report += fmt.Sprintf("  文档池分配: %d\n", documentPoolMetrics.Allocations)
	report += fmt.Sprintf("  文档池重用: %d\n", documentPoolMetrics.Reuses)
	report += fmt.Sprintf("  文档池未命中: %d\n", documentPoolMetrics.Misses)
	report += "\n"
	
	// Alerts
	alerts := poa.resourceManager.GetAlerts()
	if len(alerts) > 0 {
		report += "资源警告:\n"
		for _, alert := range alerts {
			report += fmt.Sprintf("  [%s] %s: %s\n", alert.Severity, alert.Type, alert.Message)
		}
		report += "\n"
	}
	
	return report
}

// Advanced performance utilities

// OptimizeMemoryUsage provides memory optimization utilities
func OptimizeMemoryUsage() {
	// Force garbage collection
	runtime.GC()
	
	// Set memory limit (Go 1.19+)
	// debug.SetMemoryLimit(100 * 1024 * 1024) // 100MB
	
	// Set GC percentage
	// debug.SetGCPercent(100)
}

// OptimizeConcurrency provides concurrency optimization utilities
func OptimizeConcurrency(maxProcs int) {
	if maxProcs > 0 {
		runtime.GOMAXPROCS(maxProcs)
	}
}

// GetAdvancedPerformanceTips returns advanced performance optimization tips
func GetAdvancedPerformanceTips() []string {
	return []string{
		"🚀 使用对象池复用频繁分配的对象，减少GC压力",
		"🚀 实现缓存机制减少重复计算和I/O操作",
		"🚀 使用并发管理器控制并发数量，避免资源竞争",
		"🚀 监控系统资源使用，及时释放不需要的资源",
		"🚀 使用流式处理处理大文档，避免一次性加载全部内容",
		"🚀 实现批量操作减少函数调用开销",
		"🚀 使用内存映射文件处理超大文档",
		"🚀 实现预分配策略减少动态内存分配",
		"🚀 使用压缩算法减少内存占用",
		"🚀 实现智能缓存策略，根据访问模式调整缓存大小",
	}
}

// PrintAdvancedPerformanceTips prints advanced performance optimization tips
func PrintAdvancedPerformanceTips() {
	fmt.Println("=== 高级性能优化建议 ===")
	tips := GetAdvancedPerformanceTips()
	for i, tip := range tips {
		fmt.Printf("%d. %s\n", i+1, tip)
	}
} 