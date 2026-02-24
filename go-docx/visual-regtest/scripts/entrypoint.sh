#!/usr/bin/env bash
# entrypoint.sh – orchestrates the visual regression pipeline inside Docker.
#
# The test corpus is baked into the image at /corpus.
# Two roundtrip binaries are pre-built:
#   /usr/local/bin/roundtrip-opc   (OPC layer)
#   /usr/local/bin/roundtrip-docx  (docx layer)
#
# Report output goes to /output (bind-mounted from the host).
#
# Environment variables (all have sensible defaults):
#   LAYER           – which layer to test: opc | docx  (default: opc)
#   SSIM_THRESHOLD  – SSIM pass threshold  (default: 0.98)
#   DPI             – rendering resolution  (default: 150)
#   WORKERS         – parallel workers      (default: 4)
set -euo pipefail

LAYER="${LAYER:-opc}"
THRESHOLD="${SSIM_THRESHOLD:-0.98}"
DPI="${DPI:-150}"
WORKERS="${WORKERS:-4}"

# Validate LAYER.
case "${LAYER}" in
    opc|docx) ;;
    *)
        echo "[entrypoint] ERROR: LAYER must be 'opc' or 'docx', got '${LAYER}'"
        exit 1
        ;;
esac

ROUNDTRIP_BIN="/usr/local/bin/roundtrip-${LAYER}"
LABEL=$(echo "${LAYER}" | tr '[:lower:]' '[:upper:]')

DATA="/data"
ORIG_DIR="/corpus"
RT_DIR="${DATA}/roundtrip"
WORK_DIR="${DATA}/work"
REPORT_DIR="/output"

echo "=============================================="
echo " ${LABEL} Visual Regression Test"
echo "=============================================="
echo " Layer:     ${LAYER}"
echo " Threshold: ${THRESHOLD}"
echo " DPI:       ${DPI}"
echo " Workers:   ${WORKERS}"
echo "=============================================="

NFILES=$(find "${ORIG_DIR}" -maxdepth 1 -iname '*.docx' | wc -l)
echo "[entrypoint] found ${NFILES} .docx files in corpus"

if [ "${NFILES}" -eq 0 ]; then
    echo "[entrypoint] ERROR: no .docx files found in ${ORIG_DIR}"
    echo "[entrypoint] Put your .docx files into opc-visual-regtest/test-files/ and rebuild."
    exit 1
fi

# --------------------------------------------------------------------------
# Step 1: Run roundtrip on all .docx files.
# --------------------------------------------------------------------------
mkdir -p "${RT_DIR}"
echo "[entrypoint] running ${LABEL} roundtrip …"
"${ROUNDTRIP_BIN}" --input="${ORIG_DIR}" --output="${RT_DIR}" --workers="${WORKERS}"

# --------------------------------------------------------------------------
# Step 2: SSIM comparison + report.
# --------------------------------------------------------------------------
echo "[entrypoint] running SSIM comparison …"
python3 /opt/scripts/compare_ssim.py \
    --original-dir="${ORIG_DIR}" \
    --roundtrip-dir="${RT_DIR}" \
    --work-dir="${WORK_DIR}" \
    --report="${REPORT_DIR}/index.html" \
    --threshold="${THRESHOLD}" \
    --dpi="${DPI}" \
    --workers="${WORKERS}" \
    || true  # don't fail the container; the report has the details

echo ""
echo "=============================================="
echo " Report: opc-visual-regtest/report/index.html"
echo "=============================================="
