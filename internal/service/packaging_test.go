package service_test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/vortex/docx-api/internal/service"
)

// testDocxPath returns the path to a test .docx file.
// Place any valid .docx into test/testdata/ to enable integration testing.
func testDocxPath(t *testing.T) string {
	t.Helper()
	// Walk up to find test/testdata
	candidates := []string{
		"../../test/testdata/sample.docx",
		"test/testdata/sample.docx",
	}
	for _, p := range candidates {
		if abs, err := filepath.Abs(p); err == nil {
			if _, err := os.Stat(abs); err == nil {
				return abs
			}
		}
	}
	t.Skip("no test .docx found in test/testdata/sample.docx — skipping integration test")
	return ""
}

func loadTestDocx(t *testing.T) []byte {
	t.Helper()
	path := testDocxPath(t)
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("reading test docx: %v", err)
	}
	return data
}

func TestOpen_InvalidData(t *testing.T) {
	t.Parallel()
	svc := service.NewPackagingService()
	_, err := svc.Open([]byte("not a zip"))
	if err == nil {
		t.Error("expected error for invalid data, got nil")
	}
}

func TestRoundTrip_InvalidData(t *testing.T) {
	t.Parallel()
	svc := service.NewPackagingService()
	_, err := svc.RoundTrip([]byte("not a zip"))
	if err == nil {
		t.Error("expected error for invalid data, got nil")
	}
}

func TestValidate_InvalidData(t *testing.T) {
	t.Parallel()
	svc := service.NewPackagingService()
	_, err := svc.Validate([]byte("not a zip"))
	if err == nil {
		t.Error("expected error for invalid data, got nil")
	}
}

func TestOpen_EmptySlice(t *testing.T) {
	t.Parallel()
	svc := service.NewPackagingService()
	_, err := svc.Open([]byte{})
	if err == nil {
		t.Error("expected error for empty data, got nil")
	}
}

// Integration tests — run only when a sample .docx is present.

func TestOpen_Integration(t *testing.T) {
	data := loadTestDocx(t)
	svc := service.NewPackagingService()

	info, err := svc.Open(data)
	if err != nil {
		t.Fatalf("Open failed: %v", err)
	}

	if info.PartsCount == 0 {
		t.Error("expected at least one part")
	}
}

func TestRoundTrip_Integration(t *testing.T) {
	data := loadTestDocx(t)
	svc := service.NewPackagingService()

	output, err := svc.RoundTrip(data)
	if err != nil {
		t.Fatalf("RoundTrip failed: %v", err)
	}

	if len(output) == 0 {
		t.Error("output is empty")
	}

	// Verify the output can be re-opened.
	info, err := svc.Open(output)
	if err != nil {
		t.Fatalf("re-opening round-tripped document failed: %v", err)
	}
	if info.PartsCount == 0 {
		t.Error("round-tripped document has no parts")
	}
}

func TestValidate_Integration(t *testing.T) {
	data := loadTestDocx(t)
	svc := service.NewPackagingService()

	result, err := svc.Validate(data)
	if err != nil {
		t.Fatalf("Validate failed: %v", err)
	}

	if !result.Success {
		t.Error("validation reported failure")
	}
	if result.OutputSize == 0 {
		t.Error("output size is 0")
	}
}
