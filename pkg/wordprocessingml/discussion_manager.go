package wordprocessingml

import (
	"context"
	"sync"
	"time"

	"github.com/tanqiangyes/go-word/pkg/utils"
)

// DiscussionManager 讨论管理器
type DiscussionManager struct {
	discussions  map[string]*DiscussionManagerDiscussion
	comments     map[string]*DiscussionManagerComment
	threads      map[string]*DiscussionManagerThread
	notifications map[string]*DiscussionManagerNotification
	subscribers  map[string]*DiscussionManagerSubscriber
	mu           sync.RWMutex
	logger       *utils.Logger
	config       *DiscussionManagerConfig
	revisionTracker *RevisionTracker
}

// DiscussionManagerDiscussion 讨论
type DiscussionManagerDiscussion struct {
	ID          string                    `json:"id"`
	Title       string                    `json:"title"`
	Content     string                    `json:"content"`
	Author      string                    `json:"author"`
	Status      DiscussionManagerDiscussionStatus `json:"status"`
	Type        DiscussionManagerDiscussionType `json:"type"`
	Priority    DiscussionManagerPriority `json:"priority"`
	Tags        []string                  `json:"tags"`
	ThreadID    string                    `json:"thread_id"`
	CreatedAt   time.Time                 `json:"created_at"`
	UpdatedAt   time.Time                 `json:"updated_at"`
	Metadata    map[string]interface{}    `json:"metadata"`
}

// DiscussionManagerComment 评论
type DiscussionManagerComment struct {
	ID          string                    `json:"id"`
	DiscussionID string                   `json:"discussion_id"`
	Content     string                    `json:"content"`
	Author      string                    `json:"author"`
	Status      DiscussionManagerCommentStatus `json:"status"`
	Type        DiscussionManagerCommentType `json:"type"`
	ParentID    *string                   `json:"parent_id"`
	Replies     []*DiscussionManagerComment `json:"replies"`
	CreatedAt   time.Time                 `json:"created_at"`
	UpdatedAt   time.Time                 `json:"updated_at"`
	Metadata    map[string]interface{}    `json:"metadata"`
}

// DiscussionManagerThread 讨论线程
type DiscussionManagerThread struct {
	ID          string                    `json:"id"`
	Title       string                    `json:"title"`
	Description string                    `json:"description"`
	Status      DiscussionManagerThreadStatus `json:"status"`
	Type        DiscussionManagerThreadType `json:"type"`
	Creator     string                    `json:"creator"`
	Moderators  []string                  `json:"moderators"`
	Participants []string                  `json:"participants"`
	Discussions []*DiscussionManagerDiscussion `json:"discussions"`
	CreatedAt   time.Time                 `json:"created_at"`
	UpdatedAt   time.Time                 `json:"updated_at"`
	Metadata    map[string]interface{}    `json:"metadata"`
}

// DiscussionManagerNotification 通知
type DiscussionManagerNotification struct {
	ID          string                    `json:"id"`
	UserID      string                    `json:"user_id"`
	Type        DiscussionManagerNotificationType `json:"type"`
	Title       string                    `json:"title"`
	Content     string                    `json:"content"`
	Status      DiscussionManagerNotificationStatus `json:"status"`
	Priority    DiscussionManagerPriority `json:"priority"`
	TargetID    string                    `json:"target_id"`
	TargetType  string                    `json:"target_type"`
	CreatedAt   time.Time                 `json:"created_at"`
	ReadAt      *time.Time                `json:"read_at"`
	Metadata    map[string]interface{}    `json:"metadata"`
}

// DiscussionManagerSubscriber 订阅者
type DiscussionManagerSubscriber struct {
	ID          string                    `json:"id"`
	UserID      string                    `json:"user_id"`
	ThreadID    string                    `json:"thread_id"`
	Type        DiscussionManagerSubscriptionType `json:"type"`
	Status      DiscussionManagerSubscriptionStatus `json:"status"`
	Preferences *DiscussionManagerPreferences `json:"preferences"`
	CreatedAt   time.Time                 `json:"created_at"`
	UpdatedAt   time.Time                 `json:"updated_at"`
}

// DiscussionManagerPreferences 订阅偏好
type DiscussionManagerPreferences struct {
	EmailNotifications bool `json:"email_notifications"`
	PushNotifications  bool `json:"push_notifications"`
	DigestFrequency   DiscussionManagerDigestFrequency `json:"digest_frequency"`
	PriorityFilter     DiscussionManagerPriority `json:"priority_filter"`
}

// DiscussionManagerConfig 配置
type DiscussionManagerConfig struct {
	MaxDiscussions     int           `json:"max_discussions"`
	MaxComments        int           `json:"max_comments"`
	MaxThreads         int           `json:"max_threads"`
	MaxNotifications   int           `json:"max_notifications"`
	AutoCleanup        bool          `json:"auto_cleanup"`
	CleanupInterval    time.Duration `json:"cleanup_interval"`
	NotificationRetention time.Duration `json:"notification_retention"`
	DigestInterval     time.Duration `json:"digest_interval"`
	ModerationEnabled  bool          `json:"moderation_enabled"`
	AutoArchive        bool          `json:"auto_archive"`
}

// 常量定义
const (
	// 讨论状态
	DiscussionManagerDiscussionStatusActive   DiscussionManagerDiscussionStatus = "active"
	DiscussionManagerDiscussionStatusResolved DiscussionManagerDiscussionStatus = "resolved"
	DiscussionManagerDiscussionStatusClosed   DiscussionManagerDiscussionStatus = "closed"
	DiscussionManagerDiscussionStatusArchived DiscussionManagerDiscussionStatus = "archived"

	// 讨论类型
	DiscussionManagerDiscussionTypeQuestion    DiscussionManagerDiscussionType = "question"
	DiscussionManagerDiscussionTypeSuggestion DiscussionManagerDiscussionType = "suggestion"
	DiscussionManagerDiscussionTypeBug        DiscussionManagerDiscussionType = "bug"
	DiscussionManagerDiscussionTypeFeature    DiscussionManagerDiscussionType = "feature"
	DiscussionManagerDiscussionTypeGeneral    DiscussionManagerDiscussionType = "general"

	// 评论状态
	DiscussionManagerCommentStatusActive   DiscussionManagerCommentStatus = "active"
	DiscussionManagerCommentStatusHidden   DiscussionManagerCommentStatus = "hidden"
	DiscussionManagerCommentStatusDeleted  DiscussionManagerCommentStatus = "deleted"
	DiscussionManagerCommentStatusModerated DiscussionManagerCommentStatus = "moderated"

	// 评论类型
	DiscussionManagerCommentTypeReply     DiscussionManagerCommentType = "reply"
	DiscussionManagerCommentTypeQuestion  DiscussionManagerCommentType = "question"
	DiscussionManagerCommentTypeAnswer    DiscussionManagerCommentType = "answer"
	DiscussionManagerCommentTypeFeedback  DiscussionManagerCommentType = "feedback"
	DiscussionManagerCommentTypeModeration DiscussionManagerCommentType = "moderation"

	// 线程状态
	DiscussionManagerThreadStatusActive   DiscussionManagerThreadStatus = "active"
	DiscussionManagerThreadStatusPaused  DiscussionManagerThreadStatus = "paused"
	DiscussionManagerThreadStatusClosed  DiscussionManagerThreadStatus = "closed"
	DiscussionManagerThreadStatusArchived DiscussionManagerThreadStatus = "archived"

	// 线程类型
	DiscussionManagerThreadTypeGeneral    DiscussionManagerThreadType = "general"
	DiscussionManagerThreadTypeSupport    DiscussionManagerThreadType = "support"
	DiscussionManagerThreadTypeFeedback   DiscussionManagerThreadType = "feedback"
	DiscussionManagerThreadTypeAnnouncement DiscussionManagerThreadType = "announcement"

	// 通知类型
	DiscussionManagerNotificationTypeNewDiscussion DiscussionManagerNotificationType = "new_discussion"
	DiscussionManagerNotificationTypeNewComment    DiscussionManagerNotificationType = "new_comment"
	DiscussionManagerNotificationTypeMention       DiscussionManagerNotificationType = "mention"
	DiscussionManagerNotificationTypeReply         DiscussionManagerNotificationType = "reply"
	DiscussionManagerNotificationTypeModeration    DiscussionManagerNotificationType = "moderation"

	// 通知状态
	DiscussionManagerNotificationStatusUnread DiscussionManagerNotificationStatus = "unread"
	DiscussionManagerNotificationStatusRead   DiscussionManagerNotificationStatus = "read"
	DiscussionManagerNotificationStatusArchived DiscussionManagerNotificationStatus = "archived"

	// 订阅类型
	DiscussionManagerSubscriptionTypeAll      DiscussionManagerSubscriptionType = "all"
	DiscussionManagerSubscriptionTypePriority DiscussionManagerSubscriptionType = "priority"
	DiscussionManagerSubscriptionTypeDigest   DiscussionManagerSubscriptionType = "digest"

	// 订阅状态
	DiscussionManagerSubscriptionStatusActive   DiscussionManagerSubscriptionStatus = "active"
	DiscussionManagerSubscriptionStatusPaused  DiscussionManagerSubscriptionStatus = "paused"
	DiscussionManagerSubscriptionStatusUnsubscribed DiscussionManagerSubscriptionStatus = "unsubscribed"

	// 优先级
	DiscussionManagerPriorityLow      DiscussionManagerPriority = "low"
	DiscussionManagerPriorityMedium   DiscussionManagerPriority = "medium"
	DiscussionManagerPriorityHigh     DiscussionManagerPriority = "high"
	DiscussionManagerPriorityCritical DiscussionManagerPriority = "critical"

	// 摘要频率
	DiscussionManagerDigestFrequencyNever    DiscussionManagerDigestFrequency = "never"
	DiscussionManagerDigestFrequencyDaily    DiscussionManagerDigestFrequency = "daily"
	DiscussionManagerDigestFrequencyWeekly   DiscussionManagerDigestFrequency = "weekly"
	DiscussionManagerDigestFrequencyMonthly  DiscussionManagerDigestFrequency = "monthly"
)

// 类型定义
type DiscussionManagerDiscussionStatus string
type DiscussionManagerDiscussionType string
type DiscussionManagerCommentStatus string
type DiscussionManagerCommentType string
type DiscussionManagerThreadStatus string
type DiscussionManagerThreadType string
type DiscussionManagerNotificationType string
type DiscussionManagerNotificationStatus string
type DiscussionManagerSubscriptionType string
type DiscussionManagerSubscriptionStatus string
type DiscussionManagerPriority string
type DiscussionManagerDigestFrequency string

// NewDiscussionManager 创建新的讨论管理器
func NewDiscussionManager(config *DiscussionManagerConfig, revisionTracker *RevisionTracker) *DiscussionManager {
	if config == nil {
		config = &DiscussionManagerConfig{
			MaxDiscussions:      1000,
			MaxComments:         5000,
			MaxThreads:          100,
			MaxNotifications:    1000,
			AutoCleanup:         true,
			CleanupInterval:     24 * time.Hour,
			NotificationRetention: 30 * 24 * time.Hour,
			DigestInterval:      24 * time.Hour,
			ModerationEnabled:   true,
			AutoArchive:         true,
		}
	}

	dm := &DiscussionManager{
		discussions:    make(map[string]*DiscussionManagerDiscussion),
		comments:       make(map[string]*DiscussionManagerComment),
		threads:        make(map[string]*DiscussionManagerThread),
		notifications:  make(map[string]*DiscussionManagerNotification),
		subscribers:    make(map[string]*DiscussionManagerSubscriber),
		config:         config,
		revisionTracker: revisionTracker,
		logger:         utils.NewLogger(utils.LogLevelInfo, nil),
	}

	// 启动自动清理
	if config.AutoCleanup {
		go dm.startAutoCleanup()
	}

	return dm
}

// CreateThread 创建讨论线程
func (dm *DiscussionManager) CreateThread(ctx context.Context, thread *DiscussionManagerThread) error {
	dm.mu.Lock()
	defer dm.mu.Unlock()

	// 检查线程数量限制
	if len(dm.threads) >= dm.config.MaxThreads {
		return utils.NewStructuredDocumentError(utils.ErrResourceExhausted, "已达到最大线程数量限制")
	}

	// 生成ID
	if thread.ID == "" {
		thread.ID = utils.GenerateID()
	}

	// 设置默认值
	if thread.Status == "" {
		thread.Status = DiscussionManagerThreadStatusActive
	}
	if thread.Type == "" {
		thread.Type = DiscussionManagerThreadTypeGeneral
	}
	if thread.CreatedAt.IsZero() {
		thread.CreatedAt = time.Now()
	}
	thread.UpdatedAt = time.Now()

	// 存储线程
	dm.threads[thread.ID] = thread

	dm.logger.Info("讨论线程已创建，线程ID: %s, 标题: %s, 创建者: %s, 类型: %s", thread.ID, thread.Title, thread.Creator, thread.Type)

	return nil
}

// CreateDiscussion 创建讨论
func (dm *DiscussionManager) CreateDiscussion(ctx context.Context, discussion *DiscussionManagerDiscussion) error {
	dm.mu.Lock()
	defer dm.mu.Unlock()

	// 检查讨论数量限制
	if len(dm.discussions) >= dm.config.MaxDiscussions {
		return utils.NewStructuredDocumentError(utils.ErrResourceExhausted, "已达到最大讨论数量限制")
	}

	// 验证线程存在
	if discussion.ThreadID != "" {
		if _, exists := dm.threads[discussion.ThreadID]; !exists {
			return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "讨论线程不存在")
		}
	}

	// 生成ID
	if discussion.ID == "" {
		discussion.ID = utils.GenerateID()
	}

	// 设置默认值
	if discussion.Status == "" {
		discussion.Status = DiscussionManagerDiscussionStatusActive
	}
	if discussion.Type == "" {
		discussion.Type = DiscussionManagerDiscussionTypeGeneral
	}
	if discussion.Priority == "" {
		discussion.Priority = DiscussionManagerPriorityMedium
	}
	if discussion.CreatedAt.IsZero() {
		discussion.CreatedAt = time.Now()
	}
	discussion.UpdatedAt = time.Now()

	// 存储讨论
	dm.discussions[discussion.ID] = discussion

	// 添加到线程
	if discussion.ThreadID != "" {
		if thread, exists := dm.threads[discussion.ThreadID]; exists {
			thread.Discussions = append(thread.Discussions, discussion)
			thread.UpdatedAt = time.Now()
		}
	}

	dm.logger.Info("讨论已创建，讨论ID: %s, 标题: %s, 作者: %s, 线程ID: %s", discussion.ID, discussion.Title, discussion.Author, discussion.ThreadID)

	return nil
}

// AddComment 添加评论
func (dm *DiscussionManager) AddComment(ctx context.Context, comment *DiscussionManagerComment) error {
	dm.mu.Lock()
	defer dm.mu.Unlock()

	// 检查评论数量限制
	if len(dm.comments) >= dm.config.MaxComments {
		return utils.NewStructuredDocumentError(utils.ErrResourceExhausted, "已达到最大评论数量限制")
	}

	// 验证讨论存在
	if _, exists := dm.discussions[comment.DiscussionID]; !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "讨论不存在")
	}

	// 生成ID
	if comment.ID == "" {
		comment.ID = utils.GenerateID()
	}

	// 设置默认值
	if comment.Status == "" {
		comment.Status = DiscussionManagerCommentStatusActive
	}
	if comment.Type == "" {
		comment.Type = DiscussionManagerCommentTypeReply
	}
	if comment.CreatedAt.IsZero() {
		comment.CreatedAt = time.Now()
	}
	comment.UpdatedAt = time.Now()

	// 存储评论
	dm.comments[comment.ID] = comment

	dm.logger.Info("评论已添加，评论ID: %s, 讨论ID: %s, 作者: %s, 类型: %s", comment.ID, comment.DiscussionID, comment.Author, comment.Type)

	return nil
}

// GetThreadDiscussions 获取线程讨论
func (dm *DiscussionManager) GetThreadDiscussions(threadID string, limit int) ([]*DiscussionManagerDiscussion, error) {
	dm.mu.RLock()
	defer dm.mu.RUnlock()

	thread, exists := dm.threads[threadID]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "讨论线程不存在")
	}

	discussions := make([]*DiscussionManagerDiscussion, len(thread.Discussions))
	copy(discussions, thread.Discussions)

	// 按更新时间排序（最新的在前）
	utils.SortByTimestamp(discussions, func(d *DiscussionManagerDiscussion) time.Time {
		return d.UpdatedAt
	})

	if limit > 0 && len(discussions) > limit {
		discussions = discussions[:limit]
	}

	return discussions, nil
}

// GetDiscussionComments 获取讨论评论
func (dm *DiscussionManager) GetDiscussionComments(discussionID string, limit int) ([]*DiscussionManagerComment, error) {
	dm.mu.RLock()
	defer dm.mu.RUnlock()

	comments := make([]*DiscussionManagerComment, 0)
	for _, comment := range dm.comments {
		if comment.DiscussionID == discussionID && comment.ParentID == nil {
			comments = append(comments, comment)
		}
	}

	// 按创建时间排序（最新的在前）
	utils.SortByTimestamp(comments, func(c *DiscussionManagerComment) time.Time {
		return c.CreatedAt
	})

	if limit > 0 && len(comments) > limit {
		comments = comments[:limit]
	}

	return comments, nil
}

// startAutoCleanup 启动自动清理
func (dm *DiscussionManager) startAutoCleanup() {
	ticker := time.NewTicker(dm.config.CleanupInterval)
	defer ticker.Stop()

	for range ticker.C {
		dm.cleanup()
	}
}

// cleanup 清理过期数据
func (dm *DiscussionManager) cleanup() {
	dm.mu.Lock()
	defer dm.mu.Unlock()

	now := time.Now()
	cutoff := now.Add(-dm.config.NotificationRetention)

	// 清理过期通知
	for _, notification := range dm.notifications {
		if notification.CreatedAt.Before(cutoff) && notification.Status == DiscussionManagerNotificationStatusRead {
			delete(dm.notifications, notification.ID)
		}
	}

	// 清理已归档的讨论
	if dm.config.AutoArchive {
		archiveCutoff := now.Add(-90 * 24 * time.Hour) // 90天前
		for _, discussion := range dm.discussions {
			if discussion.UpdatedAt.Before(archiveCutoff) && discussion.Status == DiscussionManagerDiscussionStatusResolved {
				discussion.Status = DiscussionManagerDiscussionStatusArchived
			}
		}
	}

	dm.logger.Info("自动清理完成，截止时间: %v", cutoff)
}

// GetStats 获取统计信息
func (dm *DiscussionManager) GetStats() map[string]interface{} {
	dm.mu.RLock()
	defer dm.mu.RUnlock()

	stats := map[string]interface{}{
		"total_discussions": len(dm.discussions),
		"total_comments":    len(dm.comments),
		"total_threads":     len(dm.threads),
		"total_notifications": len(dm.notifications),
		"total_subscribers": len(dm.subscribers),
	}

	// 按状态统计
	discussionStatusCount := make(map[DiscussionManagerDiscussionStatus]int)
	for _, discussion := range dm.discussions {
		discussionStatusCount[discussion.Status]++
	}
	stats["discussion_status_count"] = discussionStatusCount

	commentStatusCount := make(map[DiscussionManagerCommentStatus]int)
	for _, comment := range dm.comments {
		commentStatusCount[comment.Status]++
	}
	stats["comment_status_count"] = commentStatusCount

	threadStatusCount := make(map[DiscussionManagerThreadStatus]int)
	for _, thread := range dm.threads {
		threadStatusCount[thread.Status]++
	}
	stats["thread_status_count"] = threadStatusCount

	notificationStatusCount := make(map[DiscussionManagerNotificationStatus]int)
	for _, notification := range dm.notifications {
		notificationStatusCount[notification.Status]++
	}
	stats["notification_status_count"] = notificationStatusCount

	return stats
}
