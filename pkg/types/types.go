// Package types provides shared type definitions for the go-word library
package types

import (
	"encoding/xml"
	"time"
)

// Paragraph represents a paragraph in the document
type Paragraph struct {
	Text       string
	Style      string
	Runs       []Run
	HasComment bool
	CommentID  string
}

// Run represents a text run with specific formatting
type Run struct {
	Text     string
	Bold     bool
	Italic   bool
	Underline bool
	FontSize int
	FontName string
	Color    string
}

// Table represents a table in the document
type Table struct {
	Rows    []TableRow
	Columns int
}

// TableRow represents a row in a table
type TableRow struct {
	Cells []TableCell
}

// TableCell represents a cell in a table
type TableCell struct {
	Text string
}

// DocumentContent represents the content of the document
type DocumentContent struct {
	Paragraphs []Paragraph
	Tables     []Table
	Text       string
}

// 通用Word格式属性类型
type Bold struct {
	XMLName xml.Name `xml:"b"`
	Val     string   `xml:"val,attr,omitempty"`
}

type Italic struct {
	XMLName xml.Name `xml:"i"`
	Val     string   `xml:"val,attr,omitempty"`
}

type Size struct {
	XMLName xml.Name `xml:"sz"`
	Val     string   `xml:"val,attr"`
}

type Font struct {
	XMLName xml.Name `xml:"rFonts"`
	Ascii   string   `xml:"ascii,attr,omitempty"`
	HAnsi   string   `xml:"hAnsi,attr,omitempty"`
}

type Underline struct {
	XMLName xml.Name `xml:"u"`
	Val     string   `xml:"val,attr,omitempty"`
}

type Color struct {
	XMLName xml.Name `xml:"color"`
	Val     string   `xml:"val,attr,omitempty"`
}

// Style represents a document style
type Style struct {
	Name            string            `json:"name"`
	Type            StyleType         `json:"type"`
	BasedOn         string            `json:"basedOn,omitempty"`
	Next            string            `json:"next,omitempty"`
	Properties      *StyleProperties  `json:"properties,omitempty"`
	Conditional     bool              `json:"conditional,omitempty"`
	Custom          bool              `json:"custom,omitempty"`
	Default         bool              `json:"default,omitempty"`
	Hidden          bool              `json:"hidden,omitempty"`
	Locked          bool              `json:"locked,omitempty"`
	Priority        int               `json:"priority,omitempty"`
	QuickFormat     bool              `json:"quickFormat,omitempty"`
	SemiHidden      bool              `json:"semiHidden,omitempty"`
	UnhideWhenUsed  bool              `json:"unhideWhenUsed,omitempty"`
	ID              string            `json:"id,omitempty"`
	Description     string            `json:"description,omitempty"`
	Category        string            `json:"category,omitempty"`
	Aliases         []string          `json:"aliases,omitempty"`
	CreatedAt       *time.Time        `json:"createdAt,omitempty"`
	UpdatedAt       *time.Time        `json:"updatedAt,omitempty"`
}

// Clone creates a deep copy of the Style
func (s *Style) Clone() *Style {
	if s == nil {
		return nil
	}
	
	clone := &Style{
		Name:            s.Name,
		Type:            s.Type,
		BasedOn:         s.BasedOn,
		Next:            s.Next,
		Conditional:     s.Conditional,
		Custom:          s.Custom,
		Default:         s.Default,
		Hidden:          s.Hidden,
		Locked:          s.Locked,
		Priority:        s.Priority,
		QuickFormat:     s.QuickFormat,
		SemiHidden:      s.SemiHidden,
		UnhideWhenUsed:  s.UnhideWhenUsed,
		ID:              s.ID,
		Description:     s.Description,
		Category:        s.Category,
		CreatedAt:       s.CreatedAt,
		UpdatedAt:       s.UpdatedAt,
	}
	
	// Deep copy slices
	if s.Aliases != nil {
		clone.Aliases = make([]string, len(s.Aliases))
		copy(clone.Aliases, s.Aliases)
	}
	
	// Deep copy Properties
	if s.Properties != nil {
		clone.Properties = s.Properties.Clone()
	}
	
	return clone
}

// StyleType represents the type of a style
type StyleType string

const (
	StyleTypeParagraph StyleType = "paragraph"
	StyleTypeCharacter StyleType = "character"
	StyleTypeTable     StyleType = "table"
	StyleTypeList      StyleType = "list"
)

// StyleProperties represents the properties of a style
type StyleProperties struct {
	FontName        string  `json:"fontName,omitempty"`
	FontSize        int     `json:"fontSize,omitempty"`
	FontColor       string  `json:"fontColor,omitempty"`
	BackgroundColor string  `json:"backgroundColor,omitempty"`
	Bold            bool    `json:"bold,omitempty"`
	Italic          bool    `json:"italic,omitempty"`
	Underline       bool    `json:"underline,omitempty"`
	StrikeThrough   bool    `json:"strikeThrough,omitempty"`
	Alignment       string  `json:"alignment,omitempty"`
	LineSpacing     float64 `json:"lineSpacing,omitempty"`
	SpaceBefore     float64 `json:"spaceBefore,omitempty"`
	SpaceAfter      float64 `json:"spaceAfter,omitempty"`
	FirstLineIndent float64 `json:"firstLineIndent,omitempty"`
	LeftIndent      float64 `json:"leftIndent,omitempty"`
	RightIndent     float64 `json:"rightIndent,omitempty"`
	KeepLines       bool    `json:"keepLines,omitempty"`
	KeepNext        bool    `json:"keepNext,omitempty"`
	PageBreakBefore bool    `json:"pageBreakBefore,omitempty"`
	WidowControl    bool    `json:"widowControl,omitempty"`
}

// Clone creates a deep copy of StyleProperties
func (sp *StyleProperties) Clone() *StyleProperties {
	if sp == nil {
		return nil
	}
	
	return &StyleProperties{
		FontName:        sp.FontName,
		FontSize:        sp.FontSize,
		FontColor:       sp.FontColor,
		BackgroundColor: sp.BackgroundColor,
		Bold:            sp.Bold,
		Italic:          sp.Italic,
		Underline:       sp.Underline,
		StrikeThrough:   sp.StrikeThrough,
		Alignment:       sp.Alignment,
		LineSpacing:     sp.LineSpacing,
		SpaceBefore:     sp.SpaceBefore,
		SpaceAfter:      sp.SpaceAfter,
		FirstLineIndent: sp.FirstLineIndent,
		LeftIndent:      sp.LeftIndent,
		RightIndent:     sp.RightIndent,
		KeepLines:       sp.KeepLines,
		KeepNext:        sp.KeepNext,
		PageBreakBefore: sp.PageBreakBefore,
		WidowControl:    sp.WidowControl,
	}
}

// PDFExportConfig represents PDF export configuration
type PDFExportConfig struct {
	PageSize       PDFPageSize    `json:"pageSize"`
	Orientation    PDFOrientation `json:"orientation"`
	Margins        PDFMargins     `json:"margins"`
	Quality        PDFQuality     `json:"quality"`
	Compression    bool           `json:"compression"`
	ImageQuality   int            `json:"imageQuality"`
	IncludeImages  bool           `json:"includeImages"`
	IncludeTables  bool           `json:"includeTables"`
	IncludeHeaders bool           `json:"includeHeaders"`
	IncludeFooters bool           `json:"includeFooters"`
	FontEmbedding  bool           `json:"fontEmbedding"`
	DefaultFont    string         `json:"defaultFont"`
	FontSize       int            `json:"fontSize"`
	Permissions    PDFPermissions `json:"permissions"`
	Creator        string         `json:"creator"`
}

// PDFPageSize represents PDF page size
type PDFPageSize string

const (
	PDFPageSizeA3     PDFPageSize = "A3"
	PDFPageSizeA4     PDFPageSize = "A4"
	PDFPageSizeA5     PDFPageSize = "A5"
	PDFPageSizeLetter PDFPageSize = "Letter"
	PDFPageSizeLegal  PDFPageSize = "Legal"
)

// PDFOrientation represents PDF page orientation
type PDFOrientation string

const (
	PDFOrientationPortrait  PDFOrientation = "portrait"
	PDFOrientationLandscape PDFOrientation = "landscape"
)

// PDFQuality represents PDF quality level
type PDFQuality string

const (
	PDFQualityLow    PDFQuality = "low"
	PDFQualityMedium PDFQuality = "medium"
	PDFQualityHigh   PDFQuality = "high"
)

// PDFMargins represents PDF page margins
type PDFMargins struct {
	Top    float64 `json:"top"`
	Bottom float64 `json:"bottom"`
	Left   float64 `json:"left"`
	Right  float64 `json:"right"`
}

// PDFPermissions represents PDF permissions
type PDFPermissions struct {
	AllowPrint    bool `json:"allowPrint"`
	AllowCopy     bool `json:"allowCopy"`
	AllowModify   bool `json:"allowModify"`
	AllowAnnotate bool `json:"allowAnnotate"`
}

// PDFImageInfo represents PDF image information
type PDFImageInfo struct {
	ID       string  `json:"id"`
	Path     string  `json:"path"`
	Width    float64 `json:"width"`
	Height   float64 `json:"height"`
	Format   string  `json:"format,omitempty"`
	Position string  `json:"position"`
}

// DocumentProtectionConfig represents document protection configuration
type DocumentProtectionConfig struct {
	Type           ProtectionType `json:"type"`
	Enabled        bool           `json:"enabled"`
	Password       string         `json:"password,omitempty"`
	Salt           string         `json:"salt,omitempty"`
	Algorithm      string         `json:"algorithm,omitempty"`
	SpinCount      int            `json:"spinCount,omitempty"`
	Users          []string       `json:"users,omitempty"`
	Permissions    []string       `json:"permissions,omitempty"`
	ExpiryDate     *time.Time     `json:"expiryDate,omitempty"`
	ReadOnly       bool           `json:"readOnly,omitempty"`
	NoEdit         bool           `json:"noEdit,omitempty"`
	NoFormat       bool           `json:"noFormat,omitempty"`
	NoResize       bool           `json:"noResize,omitempty"`
	NoSelect       bool           `json:"noSelect,omitempty"`
	ProtectionType ProtectionType `json:"protectionType,omitempty"`
	Watermark      *WatermarkConfig `json:"watermark,omitempty"`
}

// WatermarkConfig holds watermark settings
type WatermarkConfig struct {
	Text        string  `json:"text,omitempty"`
	Font        string  `json:"font,omitempty"`
	Size        int     `json:"size,omitempty"`
	Color       string  `json:"color,omitempty"`
	Transparency float64 `json:"transparency,omitempty"`
	Rotation    float64 `json:"rotation,omitempty"`
}

// ProtectionType represents the type of document protection
type ProtectionType string

const (
	ProtectionTypeNone      ProtectionType = "none"
	ProtectionTypeReadOnly  ProtectionType = "readOnly"
	ProtectionTypeNoEdit    ProtectionType = "noEdit"
	ProtectionTypeNoFormat  ProtectionType = "noFormat"
	ProtectionTypeNoResize  ProtectionType = "noResize"
	ProtectionTypeNoSelect  ProtectionType = "noSelect"
	ProtectionTypePassword  ProtectionType = "password"
	ProtectionTypeUser      ProtectionType = "user"
)

// DocumentValidationConfig represents document validation configuration
type DocumentValidationConfig struct {
	ValidateStructure    bool     `json:"validateStructure"`
	ValidateContent      bool     `json:"validateContent"`
	ValidateStyles       bool     `json:"validateStyles"`
	ValidateLinks        bool     `json:"validateLinks"`
	ValidateImages       bool     `json:"validateImages"`
	ValidateTables       bool     `json:"validateTables"`
	ValidateHeaders      bool     `json:"validateHeaders"`
	ValidateFooters      bool     `json:"validateFooters"`
	ValidateComments     bool     `json:"validateComments"`
	ValidateRevisions    bool     `json:"validateRevisions"`
	MaxErrors            int      `json:"maxErrors"`
	StopOnFirstError     bool     `json:"stopOnFirstError"`
	CustomRules          []string `json:"customRules,omitempty"`
	ExcludeRules         []string `json:"excludeRules,omitempty"`
	Enabled              bool     `json:"enabled"`
	AutoFix              bool     `json:"autoFix"`
	StrictMode           bool     `json:"strictMode"`
}

// Image represents an image in the document
type Image struct {
	ID          string            `json:"id"`
	Path        string            `json:"path"`
	Width       float64           `json:"width"`
	Height      float64           `json:"height"`
	AltText     string            `json:"altText,omitempty"`
	Title       string            `json:"title,omitempty"`
	Format      string            `json:"format,omitempty"`
	Size        int64             `json:"size,omitempty"`
	Position    ImagePosition     `json:"position,omitempty"`
	Alignment   string            `json:"alignment,omitempty"`
	Wrapping    string            `json:"wrapping,omitempty"`
	Effects     map[string]interface{} `json:"effects,omitempty"`
	Metadata    map[string]interface{} `json:"metadata,omitempty"`
}

// ImagePosition represents the position of an image
type ImagePosition string

const (
	ImagePositionInline    ImagePosition = "inline"
	ImagePositionFloating  ImagePosition = "floating"
	ImagePositionAbsolute  ImagePosition = "absolute"
	ImagePositionRelative  ImagePosition = "relative"
)

// DocumentFormat represents the format of a document
type DocumentFormat struct {
	Type        string            `json:"type"`
	Version     string            `json:"version"`
	Encoding    string            `json:"encoding,omitempty"`
	Compression string            `json:"compression,omitempty"`
	Features    []string          `json:"features,omitempty"`
	Limitations []string          `json:"limitations,omitempty"`
	Metadata    map[string]interface{} `json:"metadata,omitempty"`
}

// CoreProperties represents the core properties of a document
type CoreProperties struct {
	Title           string     `json:"title,omitempty"`
	Subject         string     `json:"subject,omitempty"`
	Creator         string     `json:"creator,omitempty"`
	Keywords        []string   `json:"keywords,omitempty"`
	Description     string     `json:"description,omitempty"`
	Language        string     `json:"language,omitempty"`
	Category        string     `json:"category,omitempty"`
	Version         string     `json:"version,omitempty"`
	Revision        int        `json:"revision,omitempty"`
	LastModifiedBy  string     `json:"lastModifiedBy,omitempty"`
	Created         *time.Time `json:"created,omitempty"`
	Modified        *time.Time `json:"modified,omitempty"`
	LastPrinted     *time.Time `json:"lastPrinted,omitempty"`
}

// DocumentStatistics represents document statistics
type DocumentStatistics struct {
	TotalWords       int     `json:"totalWords"`
	TotalCharacters  int     `json:"totalCharacters"`
	TotalParagraphs  int     `json:"totalParagraphs"`
	TotalTables      int     `json:"totalTables"`
	TotalImages      int     `json:"totalImages"`
	TotalPages       int     `json:"totalPages"`
	TotalSections    int     `json:"totalSections"`
	TotalHeaders     int     `json:"totalHeaders"`
	TotalFooters     int     `json:"totalFooters"`
	TotalComments    int     `json:"totalComments"`
	TotalRevisions   int     `json:"totalRevisions"`
	FileSize         int64   `json:"fileSize"`
	CreationDate     *time.Time `json:"creationDate,omitempty"`
	ModificationDate *time.Time `json:"modificationDate,omitempty"`
	LastSavedBy     string  `json:"lastSavedBy,omitempty"`
	RevisionNumber  int     `json:"revisionNumber,omitempty"`
	Application     string  `json:"application,omitempty"`
	Template        string  `json:"template,omitempty"`
}

// StyleConflictType represents the type of style conflict
type StyleConflictType string

const (
	StyleConflictTypeProperty    StyleConflictType = "property"
	StyleConflictTypeInheritance StyleConflictType = "inheritance"
	StyleConflictTypePriority    StyleConflictType = "priority"
	StyleConflictTypeFormat      StyleConflictType = "format"
)

// StyleConflictStatus represents the status of a style conflict
type StyleConflictStatus string

const (
	StyleConflictStatusPending   StyleConflictStatus = "pending"
	StyleConflictStatusResolved  StyleConflictStatus = "resolved"
	StyleConflictStatusFailed    StyleConflictStatus = "failed"
	StyleConflictStatusIgnored   StyleConflictStatus = "ignored"
)

// StyleConflict represents a style conflict
type StyleConflict struct {
	StyleID                string                 `json:"styleId"`
	Type                   StyleConflictType      `json:"type"`
	Description            string                 `json:"description"`
	Severity               string                 `json:"severity"`
	Resolved               bool                   `json:"resolved"`
	Resolution             string                 `json:"resolution,omitempty"`
	ResolvedBy             string                 `json:"resolvedBy,omitempty"`
	ResolvedAt             *time.Time             `json:"resolvedAt,omitempty"`
	Priority               int                    `json:"priority,omitempty"`
	NewStyle               *Style                 `json:"newStyle,omitempty"`
	ConflictingProperties  []string              `json:"conflictingProperties,omitempty"`
	ConflictDetails        map[string]interface{} `json:"conflictDetails,omitempty"`
	
	// Additional fields for advanced style management
	ID                    string                 `json:"id,omitempty"`
	StyleName             string                 `json:"styleName,omitempty"`
	OriginalStyle         *Style                 `json:"originalStyle,omitempty"`
	OriginalPriority      int                    `json:"originalPriority,omitempty"`
	NewPriority           int                    `json:"newPriority,omitempty"`
	ResolutionDate        string                 `json:"resolutionDate,omitempty"`
	
	// Conflict resolution fields
	Status                StyleConflictStatus    `json:"status,omitempty"`
	ResolvedStyle         *Style                 `json:"resolvedStyle,omitempty"`
}