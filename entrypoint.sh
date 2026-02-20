#!/bin/sh
set -e

CompileDaemon \
    -build="go build -buildvcs=false -o /tmp/docx-api ./cmd/server" \
    -command=/tmp/docx-api \
    -directory=. \
    -pattern='(.+\.go|go\.mod)$' \
    -graceful-kill=true \
    -graceful-timeout=10