# ── Stage 1: Build ─────────────────────────────────────────────
FROM golang:1.25-alpine AS builder

RUN apk add --no-cache git ca-certificates

WORKDIR /src

# Copy go.mod/go.sum first for layer caching
COPY go.mod go.sum ./
COPY docx-go/go.mod ./docx-go/

RUN go mod download

# Copy source code
COPY . .

# Build statically linked binary
RUN CGO_ENABLED=0 GOOS=linux GOARCH=amd64 \
    go build -ldflags="-s -w" -o /bin/docx-api ./cmd/server

# ── Stage 2: Runtime ───────────────────────────────────────────
FROM alpine:3.21

RUN apk add --no-cache ca-certificates tzdata && \
    addgroup -S app && adduser -S app -G app

WORKDIR /app

COPY --from=builder /bin/docx-api .

# Create upload directory
RUN mkdir -p /tmp/docx-uploads && chown app:app /tmp/docx-uploads

USER app

EXPOSE 8080

ENV PORT=8080 \
    MAX_UPLOAD_SIZE_MB=50 \
    READ_TIMEOUT=30s \
    WRITE_TIMEOUT=60s \
    SHUTDOWN_TIMEOUT=10s

HEALTHCHECK --interval=15s --timeout=5s --start-period=5s --retries=3 \
    CMD wget -qO- http://localhost:8080/health || exit 1

ENTRYPOINT ["./docx-api"]
