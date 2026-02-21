package service

import (
	"bytes"
	"fmt"
	"io"

	"github.com/vortex/docx-api/internal/packaging"
)

// DocumentInfo holds metadata extracted after opening a document.
type DocumentInfo struct {
	// Core properties
	Title       string `json:"title,omitempty"`
	Creator     string `json:"creator,omitempty"`
	Description string `json:"description,omitempty"`

	// App properties
	Application string `json:"application,omitempty"`

	// Structure counts
	PartsCount   int      `json:"parts_count"`
	HeaderCount  int      `json:"header_count"`
	FooterCount  int      `json:"footer_count"`
	MediaFiles   []string `json:"media_files,omitempty"`
	HasStyles    bool     `json:"has_styles"`
	HasNumbering bool     `json:"has_numbering"`
	HasComments  bool     `json:"has_comments"`
	HasFootnotes bool     `json:"has_footnotes"`
	HasEndnotes  bool     `json:"has_endnotes"`
}

// PackagingService defines the interface for document packaging operations.
type PackagingService interface {
	// Open parses a .docx from raw bytes and returns document metadata.
	Open(data []byte) (*DocumentInfo, error)

	// RoundTrip opens a .docx, then immediately saves it back, returning
	// the re-packaged bytes. This is the primary packaging test: if the
	// output is a valid .docx openable by Word, packaging is correct.
	RoundTrip(data []byte) ([]byte, error)

	// Validate opens a .docx, saves it, and returns both metadata and
	// a comparison summary (original size vs output size).
	Validate(data []byte) (*ValidationResult, error)
}

// ValidationResult holds the result of a validate operation.
type ValidationResult struct {
	Info         *DocumentInfo `json:"info"`
	OriginalSize int           `json:"original_size_bytes"`
	OutputSize   int           `json:"output_size_bytes"`
	Success      bool          `json:"success"`
}

// packagingService is the concrete implementation of PackagingService.
type packagingService struct{}

// NewPackagingService creates a new PackagingService instance.
func NewPackagingService() PackagingService {
	return &packagingService{}
}

func (s *packagingService) Open(data []byte) (*DocumentInfo, error) {
	doc, err := openFromBytes(data)
	if err != nil {
		return nil, fmt.Errorf("service: open document: %w", err)
	}
	return extractInfo(doc), nil
}

func (s *packagingService) RoundTrip(data []byte) ([]byte, error) {
	doc, err := openFromBytes(data)
	if err != nil {
		return nil, fmt.Errorf("service: open document: %w", err)
	}

	var buf bytes.Buffer
	if err := doc.SaveWriter(&buf); err != nil {
		return nil, fmt.Errorf("service: save document: %w", err)
	}

	return buf.Bytes(), nil
}

func (s *packagingService) Validate(data []byte) (*ValidationResult, error) {
	info, err := s.Open(data)
	if err != nil {
		return nil, err
	}

	output, err := s.RoundTrip(data)
	if err != nil {
		return nil, err
	}

	// Verify the output can be re-opened (double round-trip).
	_, err = openFromBytes(output)
	if err != nil {
		return &ValidationResult{
			Info:         info,
			OriginalSize: len(data),
			OutputSize:   len(output),
			Success:      false,
		}, fmt.Errorf("service: re-open after save failed: %w", err)
	}

	return &ValidationResult{
		Info:         info,
		OriginalSize: len(data),
		OutputSize:   len(output),
		Success:      true,
	}, nil
}

// openFromBytes wraps packaging.OpenReader for in-memory byte slices.
func openFromBytes(data []byte) (*packaging.Document, error) {
	reader := bytes.NewReader(data)
	return packaging.OpenReader(reader, int64(len(data)))
}

// extractInfo populates a DocumentInfo from an opened Document.
func extractInfo(doc *packaging.Document) *DocumentInfo {
	info := &DocumentInfo{
		HeaderCount:  len(doc.Headers),
		FooterCount:  len(doc.Footers),
		HasStyles:    doc.Styles != nil,
		HasNumbering: doc.Numbering != nil,
		HasComments:  doc.Comments != nil,
		HasFootnotes: doc.Footnotes != nil,
		HasEndnotes:  doc.Endnotes != nil,
	}

	if doc.CoreProps != nil {
		info.Title = doc.CoreProps.Title
		info.Creator = doc.CoreProps.Creator
		info.Description = doc.CoreProps.Description
	}

	if doc.AppProps != nil {
		info.Application = doc.AppProps.Application
	}

	mediaFiles := make([]string, 0, len(doc.Media))
	for name := range doc.Media {
		mediaFiles = append(mediaFiles, name)
	}
	info.MediaFiles = mediaFiles

	// Count total parts: document + typed parts + unknown
	count := 1 // document.xml always present
	if doc.Styles != nil {
		count++
	}
	if doc.Settings != nil {
		count++
	}
	if doc.Fonts != nil {
		count++
	}
	if doc.Numbering != nil {
		count++
	}
	if doc.Footnotes != nil {
		count++
	}
	if doc.Endnotes != nil {
		count++
	}
	if doc.Comments != nil {
		count++
	}
	if len(doc.Theme) > 0 {
		count++
	}
	if len(doc.WebSettings) > 0 {
		count++
	}
	count += len(doc.Headers) + len(doc.Footers) + len(doc.Media) + len(doc.UnknownParts)
	info.PartsCount = count

	return info
}

// Ensure io is imported (used by interface consumers).
var _ io.Reader = (*bytes.Reader)(nil)
