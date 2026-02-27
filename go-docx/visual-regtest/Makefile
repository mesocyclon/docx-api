# visual-regtest/Makefile
# ============================================================================
# Visual regression testing for OPC and DOCX roundtrip layers.
#
# 1. Put your .docx files into visual-regtest/test-files/
# 2. Run:  make run              (DOCX layer, default)
#     or:  make run LAYER=opc   (opc layer)
# 3. Open: visual-regtest/report/index.html
#
# Optional overrides:
#   make run LAYER=docx SSIM_THRESHOLD=0.95 DPI=200 WORKERS=8
# ============================================================================

SHELL      := /bin/bash
MAKEFLAGS  += --no-print-directory

# Defaults
LAYER          ?= docx
SSIM_THRESHOLD ?= 0.98
DPI            ?= 150
WORKERS        ?= 4
IMAGE_NAME     ?= visual-regtest
REPORT_DIR     := $(CURDIR)/report

# Project root = one level up from this Makefile.
PROJECT_ROOT := $(abspath $(CURDIR)/..)

.PHONY: build run clean report help check-files gen-files gen-clean

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'

check-files:
	@if [ -z "$$(ls -A $(CURDIR)/test-files/*.docx 2>/dev/null)" ]; then \
		echo ""; \
		echo "  ERROR: no .docx files in visual-regtest/test-files/"; \
		echo "  Put your test documents there and re-run."; \
		echo ""; \
		exit 1; \
	fi
	@echo "[make] $$(ls $(CURDIR)/test-files/*.docx | wc -l) .docx files in test-files/"

build: check-files ## Build the Docker image (includes test files)
	docker build \
		-t $(IMAGE_NAME) \
		-f $(CURDIR)/Dockerfile \
		$(PROJECT_ROOT)

run: build ## Run the full pipeline (LAYER=opc|docx)
	@mkdir -p $(REPORT_DIR)
	docker run --rm \
		-v "$(REPORT_DIR):/output" \
		--tmpfs /data:size=4G \
		-e LAYER=$(LAYER) \
		-e SSIM_THRESHOLD=$(SSIM_THRESHOLD) \
		-e DPI=$(DPI) \
		-e WORKERS=$(WORKERS) \
		$(IMAGE_NAME)
	@echo ""
	@echo "  Report: $(REPORT_DIR)/index.html"

clean: ## Remove report and Docker image
	rm -rf $(REPORT_DIR)
	-docker rmi $(IMAGE_NAME) 2>/dev/null

GEN_OUTPUT    := $(CURDIR)/gen-files/out

gen-files: ## Generate .docx test files from scratch using the public API
	@mkdir -p $(GEN_OUTPUT)
	cd $(PROJECT_ROOT) && go run ./visual-regtest/gen-files --output $(GEN_OUTPUT)
	@echo ""
	@echo "  Generated files: $(GEN_OUTPUT)/"
	@echo "  Manifest:        $(GEN_OUTPUT)/manifest.json"

gen-clean: ## Remove generated test files
	rm -rf $(GEN_OUTPUT)

report: ## Open the HTML report (macOS / Linux)
	@if [ -f "$(REPORT_DIR)/index.html" ]; then \
		command -v xdg-open >/dev/null && xdg-open "$(REPORT_DIR)/index.html" || \
		command -v open >/dev/null && open "$(REPORT_DIR)/index.html" || \
		echo "Open $(REPORT_DIR)/index.html in your browser"; \
	else \
		echo "No report found. Run 'make run' first."; \
	fi
