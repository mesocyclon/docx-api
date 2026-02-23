# OPC Visual Regression Test

Automated visual regression testing for the `opc` roundtrip layer. Verifies that
opening a `.docx` via `opc.OpenFile` and saving it back via `SaveToFile` produces
a visually identical document.

## Quick start

```bash
# 1. Drop your .docx files into the test-files/ directory
cp /path/to/your/*.docx opc-visual-regtest/test-files/

# 2. Run (from the opc-visual-regtest/ directory)
cd opc-visual-regtest
make run

# 3. Open the report
make report
# or manually: open report/index.html
```

That's it. Two steps: drop files, `make run`.

## How it works

```
┌──────────────┐      Go OPC       ┌───────────────┐
│ original.docx │───roundtrip────▶ │ roundtrip.docx │
└──────┬────────┘                  └──────┬─────────┘
       │  LibreOffice + pdftoppm          │
       ▼                                  ▼
   page PNGs                          page PNGs
       │                                  │
       └──────────┐    ┌─────────────────┘
                  ▼    ▼
              SSIM comparison
                    │
                    ▼
             HTML report with
          thumbnails & diff maps
```

1. **Go roundtrip** — `opc.OpenFile` → `SaveToFile` (parallel, 8 workers)
2. **LibreOffice headless** — original & roundtripped `.docx` → PDF
3. **pdftoppm** — each PDF page → PNG
4. **SSIM** — per-page structural similarity (scikit-image)
5. **Report** — `report/index.html` with side-by-side thumbnails, diff heatmaps, scores

Everything runs inside a single Docker container — no local dependencies needed.

## Requirements

- Docker (with BuildKit)

## Configuration

All optional. Pass as make variables:

```bash
make run SSIM_THRESHOLD=0.95 DPI=200 WORKERS=8
```

| Variable         | Default | Description                          |
|------------------|---------|--------------------------------------|
| `SSIM_THRESHOLD` | `0.98`  | Minimum acceptable SSIM score        |
| `DPI`            | `150`   | Rendering resolution for page images |
| `WORKERS`        | `4`     | Parallel worker count                |

## Make targets

| Target   | Description                       |
|----------|-----------------------------------|
| `run`    | Build image + run full pipeline   |
| `build`  | Build Docker image only           |
| `report` | Open the HTML report in a browser |
| `clean`  | Remove report dir and Docker image|
| `help`   | Show available targets            |

## Report output

```
report/
├── index.html          # main report — open in browser
├── index.json          # machine-readable results for CI
└── images/
    └── <docx-stem>/
        ├── orig-1.png  # original page rendering
        ├── rt-1.png    # roundtripped page rendering
        └── diff-1.png  # SSIM difference heatmap
```

## CI integration

The container exits with code 0 always (report is the artifact).
Parse `report/index.json` programmatically, or check stderr for the summary line.

## File layout

```
opc-visual-regtest/
├── test-files/         ← put your .docx files here
│   └── .gitkeep
├── report/             ← generated (gitignored)
├── roundtrip/
│   └── main.go         # Go CLI: OPC open → save
├── scripts/
│   ├── entrypoint.sh   # pipeline orchestrator
│   └── compare_ssim.py # SSIM comparison + HTML report
├── Dockerfile
├── docker-compose.yml
└── Makefile
```
