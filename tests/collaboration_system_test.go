package tests

import (
	"context"
	"testing"

	"github.com/tanqiangyes/go-word/pkg/word"
)

// TestRevisionTracker 测试修订跟踪器
func TestRevisionTracker(t *testing.T) {
	// 创建修订跟踪器
	config := &word.RevisionTrackerConfig{
		MaxRevisions:      100,
		MaxComments:       50,
		MaxSuggestions:    20,
		AutoCleanup:       false,
		EnableTracking:    true,
		EnableComments:    true,
		EnableSuggestions: true,
	}

	rt := word.NewRevisionTracker(config)

	// 测试跟踪修订
	revision := &word.RevisionTrackerRevision{
		Type:        word.RevisionTrackerChangeTypeInsert,
		Content:     "测试内容",
		Author:      "test_user",
		Description: "测试修订",
	}

	err := rt.TrackRevision(context.Background(), revision)
	if err != nil {
		t.Fatalf("跟踪修订失败: %v", err)
	}

	// 验证修订已添加
	retrievedRevision, err := rt.GetRevision(revision.ID)
	if err != nil {
		t.Fatalf("获取修订失败: %v", err)
	}

	if retrievedRevision.Content != "测试内容" {
		t.Errorf("修订内容不匹配，期望: 测试内容，实际: %s", retrievedRevision.Content)
	}

	// 测试添加评论
	comment := &word.RevisionTrackerComment{
		Content: "测试评论",
		Author:  "test_user",
	}

	err = rt.AddComment(context.Background(), comment)
	if err != nil {
		t.Fatalf("添加评论失败: %v", err)
	}

	// 验证评论已添加
	retrievedComment, err := rt.GetComment(comment.ID)
	if err != nil {
		t.Fatalf("获取评论失败: %v", err)
	}

	if retrievedComment.Content != "测试评论" {
		t.Errorf("评论内容不匹配，期望: 测试评论，实际: %s", retrievedComment.Content)
	}

	// 测试获取修订历史
	history := rt.GetRevisionHistory(10)
	if len(history) == 0 {
		t.Error("修订历史为空")
	}

	// 测试获取统计信息
	stats := rt.GetStats()
	if stats["total_revisions"] == 0 {
		t.Error("修订统计信息为空")
	}
}

// TestCollaborativeEditor 测试协作编辑器
func TestCollaborativeEditor(t *testing.T) {
	// 创建修订跟踪器
	revisionTracker := word.NewRevisionTracker(nil)

	// 创建协作编辑器
	config := &word.CollaborativeEditorConfig{
		MaxSessions:       10,
		MaxOperations:     100,
		MaxUsers:          5,
		AutoCleanup:       false,
		ConflictDetection: true,
		AutoResolution:    false,
	}

	ce := word.NewCollaborativeEditor(config, revisionTracker)

	// 创建用户
	creator := &word.CollaborativeEditorUser{
		ID:     "user1",
		Name:   "测试用户1",
		Email:  "user1@test.com",
		Role:   word.CollaborativeEditorUserRoleOwner,
		Status: word.CollaborativeEditorUserStatusActive,
	}

	// 创建会话
	session, err := ce.CreateSession(context.Background(), "doc1", creator, nil)
	if err != nil {
		t.Fatalf("创建会话失败: %v", err)
	}

	if session.ID == "" {
		t.Error("会话ID为空")
	}

	// 测试应用操作
	operation := &word.CollaborativeEditorOperation{
		UserID:  "user1",
		Type:    word.CollaborativeEditorOperationTypeInsert,
		Content: "插入的内容",
		Position: &word.CollaborativeEditorPosition{
			Start: 0,
			End:   10,
		},
	}

	err = ce.ApplyOperation(context.Background(), session.ID, operation)
	if err != nil {
		t.Fatalf("应用操作失败: %v", err)
	}

	// 测试获取会话操作
	operations, err := ce.GetSessionOperations(session.ID, 10)
	if err != nil {
		t.Fatalf("获取会话操作失败: %v", err)
	}

	if len(operations) == 0 {
		t.Error("会话操作为空")
	}

	// 测试获取统计信息
	stats := ce.GetStats()
	if stats["total_sessions"] == 0 {
		t.Error("会话统计信息为空")
	}
}

// TestDiscussionManager 测试讨论管理器
func TestDiscussionManager(t *testing.T) {
	// 创建修订跟踪器
	revisionTracker := word.NewRevisionTracker(nil)

	// 创建讨论管理器
	config := &word.DiscussionManagerConfig{
		MaxDiscussions:    50,
		MaxComments:       100,
		MaxThreads:        10,
		MaxNotifications:  50,
		AutoCleanup:       false,
		ModerationEnabled: true,
		AutoArchive:       false,
	}

	dm := word.NewDiscussionManager(config, revisionTracker)

	// 创建讨论线程
	thread := &word.DiscussionManagerThread{
		Title:       "测试线程",
		Description: "这是一个测试线程",
		Creator:     "user1",
		Type:        word.DiscussionManagerThreadTypeGeneral,
	}

	err := dm.CreateThread(context.Background(), thread)
	if err != nil {
		t.Fatalf("创建讨论线程失败: %v", err)
	}

	if thread.ID == "" {
		t.Error("线程ID为空")
	}

	// 创建讨论
	discussion := &word.DiscussionManagerDiscussion{
		Title:    "测试讨论",
		Content:  "这是一个测试讨论的内容",
		Author:   "user1",
		ThreadID: thread.ID,
		Type:     word.DiscussionManagerDiscussionTypeQuestion,
		Priority: word.DiscussionManagerPriorityMedium,
	}

	err = dm.CreateDiscussion(context.Background(), discussion)
	if err != nil {
		t.Fatalf("创建讨论失败: %v", err)
	}

	if discussion.ID == "" {
		t.Error("讨论ID为空")
	}

	// 测试获取线程讨论
	discussions, err := dm.GetThreadDiscussions(thread.ID, 10)
	if err != nil {
		t.Fatalf("获取线程讨论失败: %v", err)
	}

	if len(discussions) == 0 {
		t.Error("线程讨论为空")
	}

	// 测试获取统计信息
	stats := dm.GetStats()
	if stats["total_threads"] == 0 {
		t.Error("线程统计信息为空")
	}
}
