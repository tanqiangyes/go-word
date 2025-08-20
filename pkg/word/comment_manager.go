// Package wordprocessingml provides WordprocessingML document processing functionality
package word

import (
    "fmt"
    "strings"
    "time"
)

// CommentManager manages comments in a Word document
// Based on Open-XML-SDK structure
type CommentManager struct {
    // Comments collection
    Comments []Comment

    // Comment ID counter
    nextCommentID int

    // Comment properties
    Properties CommentProperties
}

// NewCommentManager creates a new comment manager
func NewCommentManager() *CommentManager {
    return &CommentManager{
        Comments:      make([]Comment, 0),
        nextCommentID: 1,
        Properties: CommentProperties{
            Visible:    true,
            Locked:     false,
            Resolved:   false,
            ShowAuthor: true,
            ShowDate:   true,
            ShowTime:   true,
        },
    }
}

// AddComment adds a new comment to the document
// Based on Open-XML-SDK AddComment method
func (cm *CommentManager) AddComment(author, text, paragraphID, runID string, startOffset, endOffset int) (*Comment, error) {
    if author == "" {
        return nil, fmt.Errorf("author cannot be empty")
    }

    if text == "" {
        return nil, fmt.Errorf("comment text cannot be empty")
    }

    // Generate comment ID
    commentID := fmt.Sprintf("comment_%d", cm.nextCommentID)
    cm.nextCommentID++

    // Create comment
    comment := Comment{
        ID:       commentID,
        Author:   author,
        Date:     time.Now().Format("2006-01-02T15:04:05Z"),
        Text:     text,
        Initials: cm.getInitials(author),
        Index:    len(cm.Comments),
        ParentID: "",
        Formatting: CommentFormatting{
            FontName: "Calibri",
            FontSize: 10,
            Color:    "000000",
        },
    }

    // Add to collection
    cm.Comments = append(cm.Comments, comment)

    return &comment, nil
}

// AddReply adds a reply to an existing comment
func (cm *CommentManager) AddReply(parentID, author, text string) (*Comment, error) {
    // Find parent comment
    parentComment := cm.GetComment(parentID)
    if parentComment == nil {
        return nil, fmt.Errorf("parent comment not found: %s", parentID)
    }

    // Create reply comment
    reply, err := cm.AddComment(author, text, "para_1", "run_1", 0, len(text))
    if err != nil {
        return nil, err
    }

    // Set parent relationship
    reply.ParentID = parentID

    return reply, nil
}

// GetComment gets a comment by ID
func (cm *CommentManager) GetComment(id string) *Comment {
    for i := range cm.Comments {
        if cm.Comments[i].ID == id {
            return &cm.Comments[i]
        }
    }
    return nil
}

// GetCommentsByAuthor gets all comments by a specific author
func (cm *CommentManager) GetCommentsByAuthor(author string) []*Comment {
    var comments []*Comment
    for i := range cm.Comments {
        if cm.Comments[i].Author == author {
            comments = append(comments, &cm.Comments[i])
        }
    }
    return comments
}

// ResolveComment marks a comment as resolved
func (cm *CommentManager) ResolveComment(id string) error {
    comment := cm.GetComment(id)
    if comment == nil {
        return fmt.Errorf("comment not found: %s", id)
    }

    // Update the comment in the slice
    for i := range cm.Comments {
        if cm.Comments[i].ID == id {
            cm.Comments[i].Formatting.Highlight = "resolved"
            return nil
        }
    }

    return fmt.Errorf("comment not found: %s", id)
}

// DeleteComment deletes a comment
func (cm *CommentManager) DeleteComment(id string) error {
    for i := range cm.Comments {
        if cm.Comments[i].ID == id {
            // Remove from collection
            cm.Comments = append(cm.Comments[:i], cm.Comments[i+1:]...)
            return nil
        }
    }
    return fmt.Errorf("comment not found: %s", id)
}

// GetCommentCount returns the total number of comments
func (cm *CommentManager) GetCommentCount() int {
    return len(cm.Comments)
}

// GetVisibleComments returns only visible comments
func (cm *CommentManager) GetVisibleComments() []*Comment {
    var comments []*Comment
    for i := range cm.Comments {
        if cm.Comments[i].Formatting.Highlight != "hidden" {
            comments = append(comments, &cm.Comments[i])
        }
    }
    return comments
}

// GetUnresolvedComments returns only unresolved comments
func (cm *CommentManager) GetUnresolvedComments() []*Comment {
    var comments []*Comment
    for i := range cm.Comments {
        if cm.Comments[i].Formatting.Highlight != "resolved" {
            comments = append(comments, &cm.Comments[i])
        }
    }
    return comments
}

// getInitials gets the initials from the author name
// Based on Open-XML-SDK implementation
func (cm *CommentManager) getInitials(author string) string {
    parts := strings.Fields(author)
    if len(parts) == 0 {
        return ""
    }

    initials := ""
    for _, part := range parts {
        if len(part) > 0 {
            // For Chinese characters, use "U" as default
            firstChar := rune(part[0])
            if firstChar >= 0x4E00 && firstChar <= 0x9FFF {
                // Chinese character range
                initials += "U"
            } else {
                // Use first letter for other characters
                initials += strings.ToUpper(string(part[0]))
            }
        }
    }
    return initials
}

// GenerateCommentsXML generates the comments XML content
// Based on Open-XML-SDK XML generation
func (cm *CommentManager) GenerateCommentsXML() string {
    if len(cm.Comments) == 0 {
        return ""
    }

    xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`

    for _, comment := range cm.Comments {
        xml += fmt.Sprintf(`
  <w:comment w:id="%s" w:author="%s" w:date="%s" w:initials="%s">`,
            comment.ID, comment.Author, comment.Date, comment.Initials)

        // Add comment text with proper formatting
        xml += fmt.Sprintf(`
    <w:p>
      <w:pPr>
        <w:pStyle w:val="CommentText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
          <w:sz w:val="20"/>
          <w:szCs w:val="20"/>
        </w:rPr>
        <w:t xml:space="preserve">%s</w:t>
      </w:r>
    </w:p>`, comment.Text)

        xml += `
  </w:comment>`
    }

    xml += `
</w:comments>`

    return xml
}

// GenerateCommentRangeStartXML generates comment range start XML
func (cm *CommentManager) GenerateCommentRangeStartXML(commentID string) string {
    return fmt.Sprintf(`<w:commentRangeStart w:id="%s"/>`, commentID)
}

// GenerateCommentRangeEndXML generates comment range end XML
func (cm *CommentManager) GenerateCommentRangeEndXML(commentID string) string {
    return fmt.Sprintf(`<w:commentRangeEnd w:id="%s"/>`, commentID)
}

// GenerateCommentReferenceXML generates comment reference XML
func (cm *CommentManager) GenerateCommentReferenceXML(commentID string) string {
    return fmt.Sprintf(`<w:commentReference w:id="%s"/>`, commentID)
}
