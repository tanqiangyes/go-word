package wordprocessingml

import (
	"context"
	"fmt"
	"sync"
	"time"

	"github.com/tanqiangyes/go-word/pkg/utils"
)

// CollaborativeEditor 协作编辑器
type CollaborativeEditor struct {
	sessions     map[string]*CollaborativeEditorSession
	operations   map[string]*CollaborativeEditorOperation
	conflicts    map[string]*CollaborativeEditorConflict
	users        map[string]*CollaborativeEditorUser
	mu           sync.RWMutex
	logger       *utils.Logger
	config       *CollaborativeEditorConfig
	revisionTracker *RevisionTracker
}

// CollaborativeEditorSession 协作会话
type CollaborativeEditorSession struct {
	ID          string                    `json:"id"`
	DocumentID  string                    `json:"document_id"`
	Users       map[string]*CollaborativeEditorUser `json:"users"`
	Operations  []*CollaborativeEditorOperation `json:"operations"`
	Status      CollaborativeEditorSessionStatus `json:"status"`
	CreatedAt   time.Time                 `json:"created_at"`
	UpdatedAt   time.Time                 `json:"updated_at"`
	Config      *CollaborativeEditorSessionConfig `json:"config"`
}

// CollaborativeEditorOperation 操作
type CollaborativeEditorOperation struct {
	ID          string                    `json:"id"`
	SessionID   string                    `json:"session_id"`
	UserID      string                    `json:"user_id"`
	Type        CollaborativeEditorOperationType `json:"type"`
	Content     string                    `json:"content"`
	Position    *CollaborativeEditorPosition `json:"position"`
	Timestamp   time.Time                 `json:"timestamp"`
	Status      CollaborativeEditorOperationStatus `json:"status"`
	Metadata    map[string]interface{}    `json:"metadata"`
}

// CollaborativeEditorConflict 冲突
type CollaborativeEditorConflict struct {
	ID          string                    `json:"id"`
	SessionID   string                    `json:"session_id"`
	OperationIDs []string                 `json:"operation_ids"`
	Type        CollaborativeEditorConflictType `json:"type"`
	Status      CollaborativeEditorConflictStatus `json:"status"`
	Resolution  *CollaborativeEditorResolution `json:"resolution"`
	CreatedAt   time.Time                 `json:"created_at"`
	ResolvedAt  *time.Time                `json:"resolved_at"`
}

// CollaborativeEditorUser 用户
type CollaborativeEditorUser struct {
	ID          string                    `json:"id"`
	Name        string                    `json:"name"`
	Email       string                    `json:"email"`
	Role        CollaborativeEditorUserRole `json:"role"`
	Status      CollaborativeEditorUserStatus `json:"status"`
	LastActive  time.Time                 `json:"last_active"`
	Permissions []CollaborativeEditorPermission `json:"permissions"`
	Metadata    map[string]interface{}    `json:"metadata"`
}

// CollaborativeEditorPosition 位置
type CollaborativeEditorPosition struct {
	Start       int    `json:"start"`
	End         int    `json:"end"`
	Paragraph   int    `json:"paragraph"`
	Line        int    `json:"line"`
	Character   int    `json:"character"`
	ElementID   string `json:"element_id"`
	ElementType string `json:"element_type"`
}

// CollaborativeEditorResolution 冲突解决
type CollaborativeEditorResolution struct {
	Type        CollaborativeEditorResolutionType `json:"type"`
	Strategy    CollaborativeEditorResolutionStrategy `json:"strategy"`
	Winner      string                    `json:"winner"`
	Merged      string                    `json:"merged"`
	ResolvedBy  string                    `json:"resolved_by"`
	Timestamp   time.Time                 `json:"timestamp"`
	Description string                    `json:"description"`
}

// CollaborativeEditorConfig 配置
type CollaborativeEditorConfig struct {
	MaxSessions       int           `json:"max_sessions"`
	MaxOperations     int           `json:"max_operations"`
	MaxUsers          int           `json:"max_users"`
	AutoCleanup       bool          `json:"auto_cleanup"`
	CleanupInterval   time.Duration `json:"cleanup_interval"`
	ConflictDetection bool          `json:"conflict_detection"`
	AutoResolution    bool          `json:"auto_resolution"`
	SyncInterval      time.Duration `json:"sync_interval"`
}

// CollaborativeEditorSessionConfig 会话配置
type CollaborativeEditorSessionConfig struct {
	MaxUsers          int           `json:"max_users"`
	ConflictDetection bool          `json:"conflict_detection"`
	AutoResolution    bool          `json:"auto_resolution"`
	SyncInterval      time.Duration `json:"sync_interval"`
	Permissions       []CollaborativeEditorPermission `json:"permissions"`
}

// 常量定义
const (
	// 会话状态
	CollaborativeEditorSessionStatusActive   CollaborativeEditorSessionStatus = "active"
	CollaborativeEditorSessionStatusPaused  CollaborativeEditorSessionStatus = "paused"
	CollaborativeEditorSessionStatusClosed  CollaborativeEditorSessionStatus = "closed"
	CollaborativeEditorSessionStatusArchived CollaborativeEditorSessionStatus = "archived"

	// 操作类型
	CollaborativeEditorOperationTypeInsert  CollaborativeEditorOperationType = "insert"
	CollaborativeEditorOperationTypeDelete  CollaborativeEditorOperationType = "delete"
	CollaborativeEditorOperationTypeReplace CollaborativeEditorOperationType = "replace"
	CollaborativeEditorOperationTypeFormat  CollaborativeEditorOperationType = "format"
	CollaborativeEditorOperationTypeMove    CollaborativeEditorOperationType = "move"
	CollaborativeEditorOperationTypeMerge   CollaborativeEditorOperationType = "merge"

	// 操作状态
	CollaborativeEditorOperationStatusPending   CollaborativeEditorOperationStatus = "pending"
	CollaborativeEditorOperationStatusApplied   CollaborativeEditorOperationStatus = "applied"
	CollaborativeEditorOperationStatusRejected  CollaborativeEditorOperationStatus = "rejected"
	CollaborativeEditorOperationStatusConflicted CollaborativeEditorOperationStatus = "conflicted"

	// 冲突类型
	CollaborativeEditorConflictTypeContent    CollaborativeEditorConflictType = "content"
	CollaborativeEditorConflictTypePosition   CollaborativeEditorConflictType = "position"
	CollaborativeEditorConflictTypeFormat     CollaborativeEditorConflictType = "format"
	CollaborativeEditorConflictTypeStructure  CollaborativeEditorConflictType = "structure"

	// 冲突状态
	CollaborativeEditorConflictStatusPending   CollaborativeEditorConflictStatus = "pending"
	CollaborativeEditorConflictStatusResolved  CollaborativeEditorConflictStatus = "resolved"
	CollaborativeEditorConflictStatusIgnored   CollaborativeEditorConflictStatus = "ignored"

	// 用户角色
	CollaborativeEditorUserRoleOwner    CollaborativeEditorUserRole = "owner"
	CollaborativeEditorUserRoleEditor   CollaborativeEditorUserRole = "editor"
	CollaborativeEditorUserRoleViewer   CollaborativeEditorUserRole = "viewer"
	CollaborativeEditorUserRoleCommenter CollaborativeEditorUserRole = "commenter"

	// 用户状态
	CollaborativeEditorUserStatusActive   CollaborativeEditorUserStatus = "active"
	CollaborativeEditorUserStatusInactive CollaborativeEditorUserStatus = "inactive"
	CollaborativeEditorUserStatusAway     CollaborativeEditorUserStatus = "away"

	// 权限
	CollaborativeEditorPermissionRead     CollaborativeEditorPermission = "read"
	CollaborativeEditorPermissionWrite    CollaborativeEditorPermission = "write"
	CollaborativeEditorPermissionComment  CollaborativeEditorPermission = "comment"
	CollaborativeEditorPermissionManage   CollaborativeEditorPermission = "manage"
	CollaborativeEditorPermissionResolve  CollaborativeEditorPermission = "resolve"

	// 解决类型
	CollaborativeEditorResolutionTypeManual   CollaborativeEditorResolutionType = "manual"
	CollaborativeEditorResolutionTypeAutomatic CollaborativeEditorResolutionType = "automatic"
	CollaborativeEditorResolutionTypeMerge    CollaborativeEditorResolutionType = "merge"

	// 解决策略
	CollaborativeEditorResolutionStrategyLastWins    CollaborativeEditorResolutionStrategy = "last_wins"
	CollaborativeEditorResolutionStrategyFirstWins   CollaborativeEditorResolutionStrategy = "first_wins"
	CollaborativeEditorResolutionStrategyMerge       CollaborativeEditorResolutionStrategy = "merge"
	CollaborativeEditorResolutionStrategyUserChoice  CollaborativeEditorResolutionStrategy = "user_choice"
)

// 类型定义
type CollaborativeEditorSessionStatus string
type CollaborativeEditorOperationType string
type CollaborativeEditorOperationStatus string
type CollaborativeEditorConflictType string
type CollaborativeEditorConflictStatus string
type CollaborativeEditorUserRole string
type CollaborativeEditorUserStatus string
type CollaborativeEditorPermission string
type CollaborativeEditorResolutionType string
type CollaborativeEditorResolutionStrategy string

// NewCollaborativeEditor 创建新的协作编辑器
func NewCollaborativeEditor(config *CollaborativeEditorConfig, revisionTracker *RevisionTracker) *CollaborativeEditor {
	if config == nil {
		config = &CollaborativeEditorConfig{
			MaxSessions:       100,
			MaxOperations:     10000,
			MaxUsers:          50,
			AutoCleanup:       true,
			CleanupInterval:   1 * time.Hour,
			ConflictDetection: true,
			AutoResolution:    false,
			SyncInterval:      100 * time.Millisecond,
		}
	}

	ce := &CollaborativeEditor{
		sessions:        make(map[string]*CollaborativeEditorSession),
		operations:      make(map[string]*CollaborativeEditorOperation),
		conflicts:       make(map[string]*CollaborativeEditorConflict),
		users:           make(map[string]*CollaborativeEditorUser),
		config:          config,
		revisionTracker: revisionTracker,
		logger:          utils.NewLogger(utils.LogLevelInfo, nil),
	}

	// 启动自动清理
	if config.AutoCleanup {
		go ce.startAutoCleanup()
	}

	return ce
}

// CreateSession 创建协作会话
func (ce *CollaborativeEditor) CreateSession(ctx context.Context, documentID string, creator *CollaborativeEditorUser, config *CollaborativeEditorSessionConfig) (*CollaborativeEditorSession, error) {
	ce.mu.Lock()
	defer ce.mu.Unlock()

	// 检查会话数量限制
	if len(ce.sessions) >= ce.config.MaxSessions {
		return nil, utils.NewStructuredDocumentError(utils.ErrResourceExhausted, "已达到最大会话数量限制")
	}

	// 创建会话
	session := &CollaborativeEditorSession{
		ID:         utils.GenerateID(),
		DocumentID: documentID,
		Users:      make(map[string]*CollaborativeEditorUser),
		Operations: make([]*CollaborativeEditorOperation, 0),
		Status:     CollaborativeEditorSessionStatusActive,
		CreatedAt:  time.Now(),
		UpdatedAt:  time.Now(),
		Config:     config,
	}

	// 添加创建者
	if creator != nil {
		session.Users[creator.ID] = creator
		ce.users[creator.ID] = creator
	}

	// 存储会话
	ce.sessions[session.ID] = session

	ce.logger.Info("协作会话已创建", map[string]interface{}{
		"session_id":   session.ID,
		"document_id":  documentID,
		"creator_id":   creator.ID,
		"creator_name": creator.Name,
	})

	return session, nil
}

// JoinSession 加入会话
func (ce *CollaborativeEditor) JoinSession(ctx context.Context, sessionID string, user *CollaborativeEditorUser) error {
	ce.mu.Lock()
	defer ce.mu.Unlock()

	session, exists := ce.sessions[sessionID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "会话不存在")
	}

	if session.Status != CollaborativeEditorSessionStatusActive {
		return utils.NewStructuredDocumentError(utils.ErrInvalidState, "会话未激活")
	}

	// 检查用户数量限制
	if session.Config != nil && len(session.Users) >= session.Config.MaxUsers {
		return utils.NewStructuredDocumentError(utils.ErrResourceExhausted, "会话用户数量已达上限")
	}

	// 添加用户
	session.Users[user.ID] = user
	ce.users[user.ID] = user
	session.UpdatedAt = time.Now()

	ce.logger.Info("用户已加入会话", map[string]interface{}{
		"session_id": sessionID,
		"user_id":    user.ID,
		"user_name":  user.Name,
		"user_role":  user.Role,
	})

	return nil
}

// LeaveSession 离开会话
func (ce *CollaborativeEditor) LeaveSession(ctx context.Context, sessionID string, userID string) error {
	ce.mu.Lock()
	defer ce.mu.Unlock()

	session, exists := ce.sessions[sessionID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "会话不存在")
	}

	user, exists := session.Users[userID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "用户不在会话中")
	}

	// 移除用户
	delete(session.Users, userID)
	session.UpdatedAt = time.Now()

	// 如果会话为空，关闭会话
	if len(session.Users) == 0 {
		session.Status = CollaborativeEditorSessionStatusClosed
	}

	ce.logger.Info("用户已离开会话", map[string]interface{}{
		"session_id": sessionID,
		"user_id":    userID,
		"user_name":  user.Name,
	})

	return nil
}

// ApplyOperation 应用操作
func (ce *CollaborativeEditor) ApplyOperation(ctx context.Context, sessionID string, operation *CollaborativeEditorOperation) error {
	ce.mu.Lock()
	defer ce.mu.Unlock()

	session, exists := ce.sessions[sessionID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "会话不存在")
	}

	if session.Status != CollaborativeEditorSessionStatusActive {
		return utils.NewStructuredDocumentError(utils.ErrInvalidState, "会话未激活")
	}

	// 生成操作ID
	if operation.ID == "" {
		operation.ID = utils.GenerateID()
	}

	// 设置操作属性
	operation.SessionID = sessionID
	operation.Timestamp = time.Now()
	operation.Status = CollaborativeEditorOperationStatusPending

	// 检查冲突
	if ce.config.ConflictDetection {
		conflicts := ce.detectConflicts(session, operation)
		if len(conflicts) > 0 {
			operation.Status = CollaborativeEditorOperationStatusConflicted
			ce.createConflict(sessionID, operation.ID, conflicts)
		}
	}

	// 存储操作
	ce.operations[operation.ID] = operation
	session.Operations = append(session.Operations, operation)
	session.UpdatedAt = time.Now()

	// 记录修订
	if ce.revisionTracker != nil {
		revision := &RevisionTrackerRevision{
			Type:        RevisionTrackerChangeType(operation.Type),
			Content:     operation.Content,
			Author:      operation.UserID,
			Description: fmt.Sprintf("协作操作: %s", operation.Type),
			Metadata: map[string]interface{}{
				"session_id": sessionID,
				"operation_id": operation.ID,
			},
		}
		ce.revisionTracker.TrackRevision(ctx, revision)
	}

	ce.logger.Info("操作已应用", map[string]interface{}{
		"session_id":   sessionID,
		"operation_id": operation.ID,
		"user_id":      operation.UserID,
		"type":         operation.Type,
		"status":       operation.Status,
	})

	return nil
}

// GetSessionOperations 获取会话操作
func (ce *CollaborativeEditor) GetSessionOperations(sessionID string, limit int) ([]*CollaborativeEditorOperation, error) {
	ce.mu.RLock()
	defer ce.mu.RUnlock()

	session, exists := ce.sessions[sessionID]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "会话不存在")
	}

	operations := make([]*CollaborativeEditorOperation, len(session.Operations))
	copy(operations, session.Operations)

	// 按时间戳排序（最新的在前）
	utils.SortByTimestamp(operations, func(o *CollaborativeEditorOperation) time.Time {
		return o.Timestamp
	})

	if limit > 0 && len(operations) > limit {
		operations = operations[:limit]
	}

	return operations, nil
}

// GetActiveUsers 获取活跃用户
func (ce *CollaborativeEditor) GetActiveUsers(sessionID string) ([]*CollaborativeEditorUser, error) {
	ce.mu.RLock()
	defer ce.mu.RUnlock()

	session, exists := ce.sessions[sessionID]
	if !exists {
		return nil, utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "会话不存在")
	}

	users := make([]*CollaborativeEditorUser, 0, len(session.Users))
	for _, user := range session.Users {
		if user.Status == CollaborativeEditorUserStatusActive {
			users = append(users, user)
		}
	}

	return users, nil
}

// ResolveConflict 解决冲突
func (ce *CollaborativeEditor) ResolveConflict(ctx context.Context, conflictID string, resolution *CollaborativeEditorResolution) error {
	ce.mu.Lock()
	defer ce.mu.Unlock()

	conflict, exists := ce.conflicts[conflictID]
	if !exists {
		return utils.NewStructuredDocumentError(utils.ErrDocumentNotFound, "冲突不存在")
	}

	if conflict.Status != CollaborativeEditorConflictStatusPending {
		return utils.NewStructuredDocumentError(utils.ErrInvalidState, "冲突已解决")
	}

	// 设置解决信息
	conflict.Resolution = resolution
	conflict.Status = CollaborativeEditorConflictStatusResolved
	now := time.Now()
	conflict.ResolvedAt = &now

	// 更新相关操作状态
	for _, operationID := range conflict.OperationIDs {
		if operation, exists := ce.operations[operationID]; exists {
			if resolution.Winner == operationID {
				operation.Status = CollaborativeEditorOperationStatusApplied
			} else {
				operation.Status = CollaborativeEditorOperationStatusRejected
			}
		}
	}

	ce.logger.Info("冲突已解决", map[string]interface{}{
		"conflict_id": conflictID,
		"resolution_type": resolution.Type,
		"resolution_strategy": resolution.Strategy,
		"resolved_by": resolution.ResolvedBy,
	})

	return nil
}

// detectConflicts 检测冲突
func (ce *CollaborativeEditor) detectConflicts(session *CollaborativeEditorSession, operation *CollaborativeEditorOperation) []*CollaborativeEditorOperation {
	conflicts := make([]*CollaborativeEditorOperation, 0)

	// 检查时间窗口内的操作
	windowStart := operation.Timestamp.Add(-5 * time.Second)
	for _, op := range session.Operations {
		if op.Timestamp.After(windowStart) && op.ID != operation.ID {
			if ce.isConflicting(operation, op) {
				conflicts = append(conflicts, op)
			}
		}
	}

	return conflicts
}

// isConflicting 检查是否冲突
func (ce *CollaborativeEditor) isConflicting(op1, op2 *CollaborativeEditorOperation) bool {
	// 相同位置的操作可能冲突
	if op1.Position != nil && op2.Position != nil {
		if op1.Position.Start == op2.Position.Start && op1.Position.End == op2.Position.End {
			return true
		}
	}

	// 相同类型的操作可能冲突
	if op1.Type == op2.Type {
		return true
	}

	return false
}

// createConflict 创建冲突记录
func (ce *CollaborativeEditor) createConflict(sessionID string, operationID string, conflicts []*CollaborativeEditorOperation) {
	conflict := &CollaborativeEditorConflict{
		ID:          utils.GenerateID(),
		SessionID:   sessionID,
		OperationIDs: append([]string{operationID}, getOperationIDs(conflicts)...),
		Type:        CollaborativeEditorConflictTypeContent,
		Status:      CollaborativeEditorConflictStatusPending,
		CreatedAt:   time.Now(),
	}

	ce.conflicts[conflict.ID] = conflict

	ce.logger.Info("检测到冲突", map[string]interface{}{
		"conflict_id": conflict.ID,
		"session_id":  sessionID,
		"operation_id": operationID,
		"conflict_count": len(conflicts),
	})
}

// getOperationIDs 获取操作ID列表
func getOperationIDs(operations []*CollaborativeEditorOperation) []string {
	ids := make([]string, len(operations))
	for i, op := range operations {
		ids[i] = op.ID
	}
	return ids
}

// startAutoCleanup 启动自动清理
func (ce *CollaborativeEditor) startAutoCleanup() {
	ticker := time.NewTicker(ce.config.CleanupInterval)
	defer ticker.Stop()

	for range ticker.C {
		ce.cleanup()
	}
}

// cleanup 清理过期数据
func (ce *CollaborativeEditor) cleanup() {
	ce.mu.Lock()
	defer ce.mu.Unlock()

	now := time.Now()
	cutoff := now.Add(-24 * time.Hour) // 24小时前

	// 清理过期的会话
	for id, session := range ce.sessions {
		if session.UpdatedAt.Before(cutoff) && session.Status == CollaborativeEditorSessionStatusClosed {
			delete(ce.sessions, id)
		}
	}

	// 清理过期的操作
	for id, operation := range ce.operations {
		if operation.Timestamp.Before(cutoff) && operation.Status != CollaborativeEditorOperationStatusApplied {
			delete(ce.operations, id)
		}
	}

	// 清理已解决的冲突
	for id, conflict := range ce.conflicts {
		if conflict.Status == CollaborativeEditorConflictStatusResolved && conflict.ResolvedAt != nil && conflict.ResolvedAt.Before(cutoff) {
			delete(ce.conflicts, id)
		}
	}

	ce.logger.Info("自动清理完成", map[string]interface{}{
		"cutoff_time": cutoff,
	})
}

// GetStats 获取统计信息
func (ce *CollaborativeEditor) GetStats() map[string]interface{} {
	ce.mu.RLock()
	defer ce.mu.RUnlock()

	stats := map[string]interface{}{
		"total_sessions": len(ce.sessions),
		"total_operations": len(ce.operations),
		"total_conflicts": len(ce.conflicts),
		"total_users": len(ce.users),
	}

	// 按状态统计
	sessionStatusCount := make(map[CollaborativeEditorSessionStatus]int)
	for _, session := range ce.sessions {
		sessionStatusCount[session.Status]++
	}
	stats["session_status_count"] = sessionStatusCount

	operationStatusCount := make(map[CollaborativeEditorOperationStatus]int)
	for _, operation := range ce.operations {
		operationStatusCount[operation.Status]++
	}
	stats["operation_status_count"] = operationStatusCount

	conflictStatusCount := make(map[CollaborativeEditorConflictStatus]int)
	for _, conflict := range ce.conflicts {
		conflictStatusCount[conflict.Status]++
	}
	stats["conflict_status_count"] = conflictStatusCount

	return stats
}
