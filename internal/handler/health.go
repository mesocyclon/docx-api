package handler

import (
	"net/http"

	"github.com/vortex/docx-api/pkg/response"
)

// Health handles GET /health and GET /ready.
func Health(w http.ResponseWriter, _ *http.Request) {
	response.JSON(w, http.StatusOK, map[string]string{
		"status": "ok",
	})
}
