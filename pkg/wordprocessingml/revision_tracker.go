package wordprocessingml

import (
	"context"
	"fmt"
	"sync"
	"time"

	"github.com/tanqiangyes/go-word/pkg/utils"
)

// RevisionTracker 修订跟踪器
type RevisionTracker struct {
	revisions    map[string]*RevisionTrackerRevision
	comments     map[string]*RevisionTrackerComment
	suggestions  map[string]*RevisionTrackerSuggestion
	history      []*RevisionTrackerHistoryEntry
	mu           sync.RWMutex
	logger       *utils.Logger
	config       *RevisionTrackerConfig
}

// RevisionTrackerRevision 修订记录
type RevisionTrackerRevision struct {
	ID          string                    `json:"id"`
	Type        RevisionTrackerChangeType `json:"type"`
	Content     string                    `json:"content"`
	Position    *RevisionTrackerPosition  `json:"position"`
	Author      string                    `json:"author"`
	Timestamp   time.Time                 `json:"timestamp"`
	Status      RevisionTrackerStatus     `json:"status"`
	Description string                    `json:"description"`
	Metadata    map[string]interface{}    `json:"metadata"`
}

// RevisionTrackerComment 评论
type RevisionTrackerComment struct {
	ID        string                    `json:"id"`
	Content   string                    `json:"content"`
	Author    string                    `json:"author"`
	Timestamp time.Time                 `json:"timestamp"`
	Status    RevisionTrackerStatus     `json:"status"`
	Replies   []*RevisionTrackerReply   `json:"replies"`
	Position  *RevisionTrackerPosition  `json:"position"`
	Tags      []string                  `json:"tags"`
}

// RevisionTrackerSuggestion 建议
type RevisionTrackerSuggestion struct {
	ID          string                    `json:"id"`
	Type        RevisionTrackerSuggestionType `json:"type"`
	Content     string                    `json:"content"`
	Original    string                    `json:"original"`
	Author      string                    `json:"author"`
	Timestamp   time.Time                 `json:"timestamp"`
	Status      RevisionTrackerStatus     `json:"status"`
	Priority    RevisionTrackerPriority   `json:"priority"`
	Category    string                    `json:"category"`
	Description string                    `json:"description"`
}

// RevisionTrackerReply 回复
type RevisionTrackerReply struct {
	ID        string                `json:"id"`
	Content   string                `json:"content"`
	Author    string                `json:"author"`
	Timestamp time.Time             `json:"timestamp"`
	Status    RevisionTrackerStatus `json:"status"`
}

// RevisionTrackerPosition 位置信息
type RevisionTrackerPosition struct {
	Start      int    `json:"start"`
	End        int    `json:"end"`
	Paragraph  int    `json:"paragraph"`
	Line       int    `json:"line"`
	Character  int    `json:"character"`
	ElementID  string `json:"element_id"`
	ElementType string `json:"element_type"`
}

// RevisionTrackerHistoryEntry 历史记录条目
type RevisionTrackerHistoryEntry struct {
	ID          string                    `json:"id"`
	Action      RevisionTrackerAction     `json:"action"`
	Target      string                    `json:"target"`
	Author      string                    `json:"author"`
	Timestamp   time.Time                 `json:"timestamp"`
	Description string                    `json:"description"`
	Metadata    map[string]interface{}    `json:"metadata"`
}

// RevisionTrackerConfig 配置
type RevisionTrackerConfig struct {
	MaxRevisions     int           `json:"max_revisions"`
	MaxComments      int           `json:"max_comments"`
	MaxSuggestions   int           `json:"max_suggestions"`
	AutoCleanup      bool          `json:"auto_cleanup"`
	CleanupInterval  time.Duration `json:"cleanup_interval"`
	EnableTracking   bool          `json:"enable_tracking"`
	EnableComments   bool          `json:"enable_comments"`
	EnableSuggestions bool         `json:"enable_suggestions"`
}

// 常量定义
const (
	// 变更类型
	RevisionTrackerChangeTypeInsert    RevisionTrackerChangeType = "insert"
	RevisionTrackerChangeTypeDelete    RevisionTrackerChangeType = "delete"
	RevisionTrackerChangeTypeReplace   RevisionTrackerChangeType = "replace"
	RevisionTrackerChangeTypeFormat    RevisionTrackerChangeType = "format"
	RevisionTrackerChangeTypeMove      RevisionTrackerChangeType = "move"
	RevisionTrackerChangeTypeMerge     RevisionTrackerChangeType = "merge"
	RevisionTrackerChangeTypeSplit     RevisionTrackerChangeType = "split"

	// 建议类型
	RevisionTrackerSuggestionTypeSpelling    RevisionTrackerSuggestionType = "spelling"
	RevisionTrackerSuggestionTypeGrammar     RevisionTrackerSuggestionType = "grammar"
	RevisionTrackerSuggestionTypeStyle       RevisionTrackerSuggestionType = "style"
	RevisionTrackerSuggestionTypeContent     RevisionTrackerSuggestionType = "content"
	RevisionTrackerSuggestionTypeFormat      RevisionTrackerSuggestionType = "format"
	RevisionTrackerSuggestionTypeStructure   RevisionTrackerSuggestionType = "structure"

	// 状态
	RevisionTrackerStatusPending    RevisionTrackerStatus = "pending"
	RevisionTrackerStatusApproved   RevisionTrackerStatus = "approved"
	RevisionTrackerStatusRejected   RevisionTrackerStatus = "rejected"
	RevisionTrackerStatusApplied    RevisionTrackerStatus = "applied"
	RevisionTrackerStatusResolved   RevisionTrackerStatus = "resolved"
	RevisionTrackerStatusArchived   RevisionTrackerStatus = "archived"

	// 优先级
	RevisionTrackerPriorityLow      RevisionTrackerPriority = "low"
	RevisionTrackerPriorityMedium   RevisionTrackerPriority = "medium"
	RevisionTrackerPriorityHigh     RevisionTrackerPriority = "high"
	RevisionTrackerPriorityCritical RevisionTrackerPriority = "critical"

	// 操作类型
	RevisionTrackerActionCreate     RevisionTrackerAction = "create"
	RevisionTrackerActionUpdate     RevisionTrackerAction = "update"
	RevisionTrackerActionDelete     RevisionTrackerAction = "delete"
	RevisionTrackerActionApprove    RevisionTrackerAction = "approve"
	RevisionTrackerActionReject     RevisionTrackerAction = "reject"
	RevisionTrackerActionApply      RevisionTrackerAction = "apply"
	RevisionTrackerActionResolve    RevisionTrackerAction = "resolve"
	RevisionTrackerActionArchive    RevisionTrackerAction = "archive"
)

// 类型定义
type RevisionTrackerChangeType string
type RevisionTrackerSuggestionType string
type RevisionTrackerStatus string
type RevisionTrackerPriority string
type RevisionTrackerAction string

// NewRevisionTracker 创建新的修订跟踪器
func NewRevisionTracker(config *RevisionTrackerConfig) *RevisionTracker {
	if config == nil {
		config = &RevisionTrackerConfig{
			MaxRevisions:      1000,
			MaxComments:       500,
			MaxSuggestions:    200,
			AutoCleanup:       true,
			CleanupInterval:   24 * time.Hour,
			EnableTracking:    true,
			EnableComments:    true,
			EnableSuggestions: true,
		}
	}

	rt := &RevisionTracker{
		revisions:   make(map[string]*RevisionTrackerRevision),
		comments:    make(map[string]*RevisionTrackerComment),
		suggestions: make(map[string]*RevisionTrackerSuggestion),
		history:     make([]*RevisionTrackerHistoryEntry, 0),
		config:      config,
		logger:      utils.NewLogger(utils.LogLevelInfo, nil),
	}

	// 启动自动清理
	if config.AutoCleanup {
		go rt.startAutoCleanup()
	}

	return rt
}

// TrackRevision 跟踪修订
func (rt *RevisionTracker) TrackRevision(ctx context.Context, revision *RevisionTrackerRevision) error {
	if !rt.config.EnableTracking {
		return nil
	}

	rt.mu.Lock()
	defer rt.mu.Unlock()

	// 生成ID
	if revision.ID == "" {
		revision.ID = utils.GenerateID()
	}

	// 设置时间戳
	if revision.Timestamp.IsZero() {
		revision.Timestamp = time.Now()
	}

	// 设置默认状态
	if revision.Status == "" {
		revision.Status = RevisionTrackerStatusPending
	}

	// 存储修订
	rt.revisions[revision.ID] = revision

	// 记录历史
	rt.addHistoryEntry(&RevisionTrackerHistoryEntry{
		ID:          utils.GenerateID(),
		Action:      RevisionTrackerActionCreate,
		Target:      revision.ID,
		Author:      revision.Author,
		Timestamp:   time.Now(),
		Description: fmt.Sprintf("创建修订: %s", revision.Description),
		Metadata: map[string]interface{}{
			"revision_type": revision.Type,
			"content":       revision.Content,
		},
	})

	rt.logger.Info("修订已跟踪", map[string]interface{}{
		"revision_id": revision.ID,
		"author":      revision.Author,
		"type":        revision.Type,
	})

	return nil
}

// AddComment 添加评论
func (rt *RevisionTracker) AddComment(ctx context.Context, comment *RevisionTrackerComment) error {
	if !rt.config.EnableComments {
		return nil
	}

	rt.mu.Lock()
	defer rt.mu.Unlock()

	// 生成ID
	if comment.ID == "" {
		comment.ID = utils.GenerateID()
	}

	// 设置时间戳
	if comment.Timestamp.IsZero() {
		comment.Timestamp = time.Now()
	}

	// 设置默认状态
	if comment.Status == "" {
		comment.Status = RevisionTrackerStatusPending
	}

	// 存储评论
	rt.comments[comment.ID] = comment

	// 记录历史
	rt.addHistoryEntry(&RevisionTrackerHistoryEntry{
		ID:          utils.GenerateID(),
		Action:      RevisionTrackerActionCreate,
		Target:      comment.ID,
		Author:      comment.Author,
		Timestamp:   time.Now(),
		Description: fmt.Sprintf("添加评论: %s", comment.Content),
		Metadata: map[string]interface{}{
			"content": comment.Content,
		},
	})

	rt.logger.Info("评论已添加", map[string]interface{}{
		"comment_id": comment.ID,
		"author":     comment.Author,
	})

	return nil
}

// AddSuggestion 添加建议
func (rt *RevisionTracker) AddSuggestion(ctx context.Context, suggestion *RevisionTrackerSuggestion) error {
	if !rt.config.EnableSuggestions {
		return nil
	}

	rt.mu.Lock()
	defer rt.mu.Unlock()

	// 生成ID
	if suggestion.ID == "" {
		suggestion.ID = utils.GenerateID()
	}

	// 设置时间戳
	if suggestion.Timestamp.IsZero() {
		suggestion.Timestamp = time.Now()
	}

	// 设置默认状态
	if suggestion.Status == "" {
		suggestion.Status = RevisionTrackerStatusPending
	}

	// 设置默认优先级
	if suggestion.Priority == "" {
		suggestion.Priority = RevisionTrackerPriorityMedium
	}

	// 存储建议
	rt.suggestions[suggestion.ID] = suggestion

	// 记录历史
	rt.addHistoryEntry(&RevisionTrackerHistoryEntry{
		ID:          utils.GenerateID(),
		Action:      RevisionTrackerActionCreate,
		Target:      suggestion.ID,
		Author:      suggestion.Author,
		Timestamp:   time.Now(),
		Description: fmt.Sprintf("添加建议: %s", suggestion.Description),
		Metadata: map[string]interface{}{
			"suggestion_type": suggestion.Type,
			"priority":        suggestion.Priority,
			"category":        suggestion.Category,
		},
	})

	rt.logger.Info("建议已添加", map[string]interface{}{
		"suggestion_id": suggestion.ID,
		"author":        suggestion.Author,
		"type":          suggestion.Type,
		"priority":      suggestion.Priority,
	})

	return nil
}

// GetRevision 获取修订
func (rt *RevisionTracker) GetRevision(id string) (*RevisionTrackerRevision, error) {
	rt.mu.RLock()
	defer rt.mu.RUnlock()

	revision, exists := rt.revisions[id]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "修订不存在")
	}

	return revision, nil
}

// GetComment 获取评论
func (rt *RevisionTracker) GetComment(id string) (*RevisionTrackerComment, error) {
	rt.mu.RLock()
	defer rt.mu.RUnlock()

	comment, exists := rt.comments[id]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "评论不存在")
	}

	return comment, nil
}

// GetSuggestion 获取建议
func (rt *RevisionTracker) GetSuggestion(id string) (*RevisionTrackerSuggestion, error) {
	rt.mu.RLock()
	defer rt.mu.RUnlock()

	suggestion, exists := rt.suggestions[id]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "建议不存在")
	}

	return suggestion, nil
}

// UpdateRevisionStatus 更新修订状态
func (rt *RevisionTracker) UpdateRevisionStatus(id string, status RevisionTrackerStatus, author string) error {
	rt.mu.Lock()
	defer rt.mu.Unlock()

	revision, exists := rt.revisions[id]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "修订不存在")
	}

	oldStatus := revision.Status
	revision.Status = status

	// 记录历史
	rt.addHistoryEntry(&RevisionTrackerHistoryEntry{
		ID:          utils.GenerateID(),
		Action:      RevisionTrackerActionUpdate,
		Target:      id,
		Author:      author,
		Timestamp:   time.Now(),
		Description: fmt.Sprintf("更新修订状态: %s -> %s", oldStatus, status),
		Metadata: map[string]interface{}{
			"old_status": oldStatus,
			"new_status": status,
		},
	})

	rt.logger.Info("修订状态已更新", map[string]interface{}{
		"revision_id": id,
		"old_status":  oldStatus,
		"new_status":  status,
		"author":      author,
	})

	return nil
}

// GetRevisionHistory 获取修订历史
func (rt *RevisionTracker) GetRevisionHistory(limit int) []*RevisionTrackerRevision {
	rt.mu.RLock()
	defer rt.mu.RUnlock()

	revisions := make([]*RevisionTrackerRevision, 0, len(rt.revisions))
	for _, revision := range rt.revisions {
		revisions = append(revisions, revision)
	}

	// 按时间戳排序（最新的在前）
	utils.SortByTimestamp(revisions, func(r *RevisionTrackerRevision) time.Time {
		return r.Timestamp
	})

	if limit > 0 && len(revisions) > limit {
		revisions = revisions[:limit]
	}

	return revisions
}

// GetCommentsByAuthor 按作者获取评论
func (rt *RevisionTracker) GetCommentsByAuthor(author string) []*RevisionTrackerComment {
	rt.mu.RLock()
	defer rt.mu.RUnlock()

	comments := make([]*RevisionTrackerComment, 0)
	for _, comment := range rt.comments {
		if comment.Author == author {
			comments = append(comments, comment)
		}
	}

	// 按时间戳排序（最新的在前）
	utils.SortByTimestamp(comments, func(c *RevisionTrackerComment) time.Time {
		return c.Timestamp
	})

	return comments
}

// GetSuggestionsByType 按类型获取建议
func (rt *RevisionTracker) GetSuggestionsByType(suggestionType RevisionTrackerSuggestionType) []*RevisionTrackerSuggestion {
	rt.mu.RLock()
	defer rt.mu.RUnlock()

	suggestions := make([]*RevisionTrackerSuggestion, 0)
	for _, suggestion := range rt.suggestions {
		if suggestion.Type == suggestionType {
			suggestions = append(suggestions, suggestion)
		}
	}

	// 按优先级和时间戳排序
	utils.SortByPriorityAndTimestamp(suggestions, func(s *RevisionTrackerSuggestion) (string, time.Time) {
		return string(s.Priority), s.Timestamp
	})

	return suggestions
}

// MergeRevisions 合并修订
func (rt *RevisionTracker) MergeRevisions(revisionIDs []string, author string) (*RevisionTrackerRevision, error) {
	rt.mu.Lock()
	defer rt.mu.Unlock()

	if len(revisionIDs) < 2 {
		return nil, utils.NewStructuredDocumentError(utils.ErrInvalidInput, "至少需要两个修订进行合并")
	}

	// 获取所有修订
	revisions := make([]*RevisionTrackerRevision, 0, len(revisionIDs))
	for _, id := range revisionIDs {
		revision, exists := rt.revisions[id]
		if !exists {
			return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, fmt.Sprintf("修订不存在: %s", id))
		}
		revisions = append(revisions, revision)
	}

	// 创建合并修订
	mergedRevision := &RevisionTrackerRevision{
		ID:          utils.GenerateID(),
		Type:        RevisionTrackerChangeTypeMerge,
		Content:     "", // 合并后的内容
		Author:      author,
		Timestamp:   time.Now(),
		Status:      RevisionTrackerStatusApplied,
		Description: fmt.Sprintf("合并 %d 个修订", len(revisionIDs)),
		Metadata: map[string]interface{}{
			"merged_revisions": revisionIDs,
			"merge_count":      len(revisionIDs),
		},
	}

	// 存储合并修订
	rt.revisions[mergedRevision.ID] = mergedRevision

	// 更新原修订状态
	for _, id := range revisionIDs {
		if revision, exists := rt.revisions[id]; exists {
			revision.Status = RevisionTrackerStatusApplied
		}
	}

	// 记录历史
	rt.addHistoryEntry(&RevisionTrackerHistoryEntry{
		ID:          utils.GenerateID(),
		Action:      RevisionTrackerActionApply,
		Target:      mergedRevision.ID,
		Author:      author,
		Timestamp:   time.Now(),
		Description: fmt.Sprintf("合并修订: %s", mergedRevision.Description),
		Metadata: map[string]interface{}{
			"merged_revisions": revisionIDs,
		},
	})

	rt.logger.Info("修订已合并", map[string]interface{}{
		"merged_revision_id": mergedRevision.ID,
		"revision_count":     len(revisionIDs),
		"author":             author,
	})

	return mergedRevision, nil
}

// addHistoryEntry 添加历史记录条目
func (rt *RevisionTracker) addHistoryEntry(entry *RevisionTrackerHistoryEntry) {
	rt.history = append(rt.history, entry)

	// 限制历史记录数量
	if len(rt.history) > rt.config.MaxRevisions {
		rt.history = rt.history[1:]
	}
}

// startAutoCleanup 启动自动清理
func (rt *RevisionTracker) startAutoCleanup() {
	ticker := time.NewTicker(rt.config.CleanupInterval)
	defer ticker.Stop()

	for range ticker.C {
		rt.cleanup()
	}
}

// cleanup 清理过期数据
func (rt *RevisionTracker) cleanup() {
	rt.mu.Lock()
	defer rt.mu.Unlock()

	now := time.Now()
	cutoff := now.Add(-30 * 24 * time.Hour) // 30天前

	// 清理过期的修订
	for id, revision := range rt.revisions {
		if revision.Timestamp.Before(cutoff) && revision.Status == RevisionTrackerStatusArchived {
			delete(rt.revisions, id)
		}
	}

	// 清理过期的评论
	for id, comment := range rt.comments {
		if comment.Timestamp.Before(cutoff) && comment.Status == RevisionTrackerStatusResolved {
			delete(rt.comments, id)
		}
	}

	// 清理过期的建议
	for id, suggestion := range rt.suggestions {
		if suggestion.Timestamp.Before(cutoff) && suggestion.Status == RevisionTrackerStatusApplied {
			delete(rt.suggestions, id)
		}
	}

	rt.logger.Info("自动清理完成", map[string]interface{}{
		"cutoff_time": cutoff,
	})
}

// GetStats 获取统计信息
func (rt *RevisionTracker) GetStats() map[string]interface{} {
	rt.mu.RLock()
	defer rt.mu.RUnlock()

	stats := map[string]interface{}{
		"total_revisions": len(rt.revisions),
		"total_comments":  len(rt.comments),
		"total_suggestions": len(rt.suggestions),
		"total_history":   len(rt.history),
	}

	// 按状态统计
	revisionStatusCount := make(map[RevisionTrackerStatus]int)
	for _, revision := range rt.revisions {
		revisionStatusCount[revision.Status]++
	}
	stats["revision_status_count"] = revisionStatusCount

	commentStatusCount := make(map[RevisionTrackerStatus]int)
	for _, comment := range rt.comments {
		commentStatusCount[comment.Status]++
	}
	stats["comment_status_count"] = commentStatusCount

	suggestionStatusCount := make(map[RevisionTrackerStatus]int)
	for _, suggestion := range rt.suggestions {
		suggestionStatusCount[suggestion.Status]++
	}
	stats["suggestion_status_count"] = suggestionStatusCount

	return stats
}
