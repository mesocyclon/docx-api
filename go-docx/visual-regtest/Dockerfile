# ===========================================================================
# Visual Regression Test – self-contained Docker image
#
# Contains: Go 1.25 · LibreOffice (headless) · poppler-utils · Python 3 +
#           scikit-image, Pillow, numpy
#
# Test corpus is baked into the image from visual-regtest/test-files/.
#
# Two roundtrip binaries are built:
#   roundtrip-opc  – OPC layer only  (opc.OpenFile → SaveToFile)
#   roundtrip-docx – full docx layer (docx.OpenFile → doc.SaveFile)
#
# Select which layer to test via the LAYER env variable (default: opc).
# ===========================================================================
FROM golang:1.25-bookworm

# ---------- system dependencies -------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        poppler-utils \
        python3 \
        python3-pip \
        python3-venv \
        fonts-liberation \
        fonts-dejavu \
        fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

# ---------- Python dependencies -------------------------------------------
RUN python3 -m pip install --break-system-packages --no-cache-dir \
        scikit-image \
        Pillow \
        numpy

# ---------- project source & build ----------------------------------------
WORKDIR /workspace
COPY . .
RUN go build -o /usr/local/bin/roundtrip-opc  ./visual-regtest/roundtrip/opc \
 && go build -o /usr/local/bin/roundtrip-docx ./visual-regtest/roundtrip/docx

# ---------- test corpus (baked in) ----------------------------------------
COPY visual-regtest/test-files/ /corpus/

# ---------- scripts -------------------------------------------------------
COPY visual-regtest/scripts/ /opt/scripts/
RUN chmod +x /opt/scripts/entrypoint.sh

# ---------- runtime -------------------------------------------------------
ENV LAYER=docx \
    SSIM_THRESHOLD=0.98 \
    DPI=150 \
    WORKERS=4

ENTRYPOINT ["/opt/scripts/entrypoint.sh"]
