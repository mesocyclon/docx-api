package handler_test

import (
	"bytes"
	"encoding/json"
	"io"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"testing"

	"github.com/vortex/docx-api/internal/handler"
	"github.com/vortex/docx-api/internal/service"
)

// mockService implements service.PackagingService for testing handlers.
type mockService struct {
	openFn     func([]byte) (*service.DocumentInfo, error)
	roundTrip  func([]byte) ([]byte, error)
	validateFn func([]byte) (*service.ValidationResult, error)
}

func (m *mockService) Open(data []byte) (*service.DocumentInfo, error) {
	if m.openFn != nil {
		return m.openFn(data)
	}
	return &service.DocumentInfo{PartsCount: 5, HasStyles: true}, nil
}

func (m *mockService) RoundTrip(data []byte) ([]byte, error) {
	if m.roundTrip != nil {
		return m.roundTrip(data)
	}
	return data, nil
}

func (m *mockService) Validate(data []byte) (*service.ValidationResult, error) {
	if m.validateFn != nil {
		return m.validateFn(data)
	}
	return &service.ValidationResult{
		Info:         &service.DocumentInfo{PartsCount: 5},
		OriginalSize: len(data),
		OutputSize:   len(data),
		Success:      true,
	}, nil
}

func newMultipartRequest(t *testing.T, url string, fileData []byte) *http.Request {
	t.Helper()
	var buf bytes.Buffer
	w := multipart.NewWriter(&buf)
	fw, err := w.CreateFormFile("file", "test.docx")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := fw.Write(fileData); err != nil {
		t.Fatal(err)
	}
	w.Close()

	req := httptest.NewRequest(http.MethodPost, url, &buf)
	req.Header.Set("Content-Type", w.FormDataContentType())
	return req
}

func TestHealth(t *testing.T) {
	t.Parallel()
	rec := httptest.NewRecorder()
	req := httptest.NewRequest(http.MethodGet, "/health", nil)

	handler.Health(rec, req)

	if rec.Code != http.StatusOK {
		t.Errorf("expected 200, got %d", rec.Code)
	}

	var body map[string]string
	if err := json.NewDecoder(rec.Body).Decode(&body); err != nil {
		t.Fatal(err)
	}
	if body["status"] != "ok" {
		t.Errorf("expected status=ok, got %s", body["status"])
	}
}

func TestOpenHandler_Success(t *testing.T) {
	t.Parallel()
	svc := &mockService{}
	h := handler.NewPackagingHandler(svc)

	req := newMultipartRequest(t, "/api/v1/documents/open", []byte("fake-docx"))
	rec := httptest.NewRecorder()

	h.Open(rec, req)

	if rec.Code != http.StatusOK {
		t.Errorf("expected 200, got %d", rec.Code)
	}

	var info service.DocumentInfo
	if err := json.NewDecoder(rec.Body).Decode(&info); err != nil {
		t.Fatal(err)
	}
	if info.PartsCount != 5 {
		t.Errorf("expected 5 parts, got %d", info.PartsCount)
	}
}

func TestOpenHandler_NoFile(t *testing.T) {
	t.Parallel()
	svc := &mockService{}
	h := handler.NewPackagingHandler(svc)

	req := httptest.NewRequest(http.MethodPost, "/api/v1/documents/open", nil)
	req.Header.Set("Content-Type", "multipart/form-data")
	rec := httptest.NewRecorder()

	h.Open(rec, req)

	if rec.Code != http.StatusBadRequest {
		t.Errorf("expected 400, got %d", rec.Code)
	}
}

func TestRoundTripHandler_ReturnsDocx(t *testing.T) {
	t.Parallel()
	testData := []byte("fake-docx-bytes")
	svc := &mockService{
		roundTrip: func(data []byte) ([]byte, error) {
			return data, nil
		},
	}
	h := handler.NewPackagingHandler(svc)

	req := newMultipartRequest(t, "/api/v1/documents/roundtrip", testData)
	rec := httptest.NewRecorder()

	h.RoundTrip(rec, req)

	if rec.Code != http.StatusOK {
		t.Errorf("expected 200, got %d", rec.Code)
	}

	ct := rec.Header().Get("Content-Type")
	expected := "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
	if ct != expected {
		t.Errorf("expected content-type %s, got %s", expected, ct)
	}

	body, _ := io.ReadAll(rec.Body)
	if !bytes.Equal(body, testData) {
		t.Error("response body doesn't match input")
	}
}

func TestValidateHandler_Success(t *testing.T) {
	t.Parallel()
	svc := &mockService{}
	h := handler.NewPackagingHandler(svc)

	req := newMultipartRequest(t, "/api/v1/documents/validate", []byte("fake"))
	rec := httptest.NewRecorder()

	h.Validate(rec, req)

	if rec.Code != http.StatusOK {
		t.Errorf("expected 200, got %d", rec.Code)
	}

	var result service.ValidationResult
	if err := json.NewDecoder(rec.Body).Decode(&result); err != nil {
		t.Fatal(err)
	}
	if !result.Success {
		t.Error("expected success=true")
	}
}
