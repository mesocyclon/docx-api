package handler

import (
	"io"
	"net/http"

	"github.com/vortex/docx-api/internal/service"
	"github.com/vortex/docx-api/pkg/response"
)

// PackagingHandler exposes HTTP endpoints for testing docx packaging.
type PackagingHandler struct {
	svc service.PackagingService
}

// NewPackagingHandler creates a handler backed by the given service.
func NewPackagingHandler(svc service.PackagingService) *PackagingHandler {
	return &PackagingHandler{svc: svc}
}

// Open handles POST /api/v1/documents/open
// Accepts a multipart form with a "file" field containing a .docx.
// Returns JSON metadata about the document.
func (h *PackagingHandler) Open(w http.ResponseWriter, r *http.Request) {
	data, err := readUploadedFile(r)
	if err != nil {
		response.Error(w, http.StatusBadRequest, err.Error())
		return
	}

	info, err := h.svc.Open(data)
	if err != nil {
		response.Error(w, http.StatusUnprocessableEntity, err.Error())
		return
	}

	response.JSON(w, http.StatusOK, info)
}

// RoundTrip handles POST /api/v1/documents/roundtrip
// Accepts a .docx, opens it, re-saves it, and returns the new .docx file.
func (h *PackagingHandler) RoundTrip(w http.ResponseWriter, r *http.Request) {
	data, err := readUploadedFile(r)
	if err != nil {
		response.Error(w, http.StatusBadRequest, err.Error())
		return
	}

	output, err := h.svc.RoundTrip(data)
	if err != nil {
		response.Error(w, http.StatusUnprocessableEntity, err.Error())
		return
	}

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
	w.Header().Set("Content-Disposition", `attachment; filename="roundtrip.docx"`)
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(output)
}

// Validate handles POST /api/v1/documents/validate
// Opens the document, round-trips it, then re-opens the output to verify
// packaging integrity. Returns a JSON validation report.
func (h *PackagingHandler) Validate(w http.ResponseWriter, r *http.Request) {
	data, err := readUploadedFile(r)
	if err != nil {
		response.Error(w, http.StatusBadRequest, err.Error())
		return
	}

	result, err := h.svc.Validate(data)
	if err != nil {
		// If validation itself failed but we still have partial info, return it.
		if result != nil {
			response.JSON(w, http.StatusUnprocessableEntity, map[string]any{
				"result": result,
				"error":  err.Error(),
			})
			return
		}
		response.Error(w, http.StatusUnprocessableEntity, err.Error())
		return
	}

	response.JSON(w, http.StatusOK, result)
}

// readUploadedFile extracts the file bytes from a multipart upload.
// It looks for a form field named "file".
func readUploadedFile(r *http.Request) ([]byte, error) {
	if err := r.ParseMultipartForm(100 << 20); err != nil { // 100 MB max
		return nil, err
	}

	file, _, err := r.FormFile("file")
	if err != nil {
		return nil, err
	}
	defer file.Close()

	return io.ReadAll(file)
}
