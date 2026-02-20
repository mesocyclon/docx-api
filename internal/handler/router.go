package handler

import (
	"log/slog"
	"net/http"

	"github.com/vortex/docx-api/internal/middleware"
	"github.com/vortex/docx-api/internal/service"
)

// NewRouter builds the HTTP mux with all routes and middleware.
func NewRouter(logger *slog.Logger, svc service.PackagingService, maxBodyBytes int64) http.Handler {
	mux := http.NewServeMux()

	pkg := NewPackagingHandler(svc)

	// Health endpoints
	mux.HandleFunc("GET /health", Health)
	mux.HandleFunc("GET /ready", Health)

	// Packaging test endpoints
	mux.HandleFunc("POST /api/v1/documents/open", pkg.Open)
	mux.HandleFunc("POST /api/v1/documents/roundtrip", pkg.RoundTrip)
	mux.HandleFunc("POST /api/v1/documents/validate", pkg.Validate)

	// Apply middleware chain (outermost first)
	var h http.Handler = mux
	h = middleware.MaxBodySize(maxBodyBytes)(h)
	h = middleware.CORS(h)
	h = middleware.Recovery(logger)(h)
	h = middleware.Logging(logger)(h)

	return h
}
