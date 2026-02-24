// roundtrip reads every .docx from --input, opens it via the top-level
// docx layer (docx.OpenFile → doc.SaveFile), re-serialises it unchanged,
// and writes the result to --output.
//
// This is the docx-layer counterpart of ../opc/main.go which tests only
// the OPC packaging layer. This binary exercises the full stack:
// OPC → XML parsing → typed parts (DocumentPart, StylesPart, …) →
// re-serialisation → OPC repacking.
//
// Exit code 0  = all files processed (some may have had errors).
// A per-file JSON manifest is written to --output/manifest.json so
// downstream tools know which files succeeded and which failed.
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
)

// FileResult captures the outcome of one roundtrip.
type FileResult struct {
	Name    string `json:"name"`
	OK      bool   `json:"ok"`
	Error   string `json:"error,omitempty"`
	Elapsed string `json:"elapsed"`
}

func main() {
	inputDir := flag.String("input", "", "directory containing original .docx files")
	outputDir := flag.String("output", "", "directory for roundtripped .docx files")
	workers := flag.Int("workers", 8, "parallel workers")
	flag.Parse()

	if *inputDir == "" || *outputDir == "" {
		log.Fatal("--input and --output are required")
	}

	if err := os.MkdirAll(*outputDir, 0o755); err != nil {
		log.Fatalf("creating output dir: %v", err)
	}

	// Collect .docx paths.
	var files []string
	entries, err := os.ReadDir(*inputDir)
	if err != nil {
		log.Fatalf("reading input dir: %v", err)
	}
	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		lower := strings.ToLower(e.Name())
		if strings.HasSuffix(lower, ".docx") {
			files = append(files, e.Name())
		}
	}
	log.Printf("found %d .docx files", len(files))

	// Process in parallel.
	type job struct{ name string }
	jobs := make(chan job, len(files))
	for _, f := range files {
		jobs <- job{name: f}
	}
	close(jobs)

	var (
		mu      sync.Mutex
		results []FileResult
	)

	var wg sync.WaitGroup
	for i := 0; i < *workers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for j := range jobs {
				r := processFile(j.name, *inputDir, *outputDir)
				mu.Lock()
				results = append(results, r)
				mu.Unlock()
				if !r.OK {
					log.Printf("FAIL %s: %s", j.name, r.Error)
				}
			}
		}()
	}
	wg.Wait()

	// Write manifest.
	manifestPath := filepath.Join(*outputDir, "manifest.json")
	data, _ := json.MarshalIndent(results, "", "  ")
	if err := os.WriteFile(manifestPath, data, 0o644); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}

	okCount := 0
	for _, r := range results {
		if r.OK {
			okCount++
		}
	}
	log.Printf("done: %d/%d succeeded", okCount, len(results))
}

func processFile(name, inputDir, outputDir string) FileResult {
	start := time.Now()
	srcPath := filepath.Join(inputDir, name)
	dstPath := filepath.Join(outputDir, name)

	doc, err := docx.OpenFile(srcPath)
	if err != nil {
		return FileResult{Name: name, OK: false, Error: fmt.Sprintf("open: %v", err), Elapsed: time.Since(start).String()}
	}

	if err := doc.SaveFile(dstPath); err != nil {
		return FileResult{Name: name, OK: false, Error: fmt.Sprintf("save: %v", err), Elapsed: time.Since(start).String()}
	}

	return FileResult{Name: name, OK: true, Elapsed: time.Since(start).String()}
}
