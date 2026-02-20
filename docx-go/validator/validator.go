// Package validator checks a packaging.Document for structural issues that
// would cause MS Word to repair the file on open.
//
// Contract: C-31 in contracts.md
// Primary import: packaging (transitive access to wml/* types for deep checks)
package validator

import (
	"fmt"
	"strings"

	"github.com/vortex/docx-go/packaging"
	"github.com/vortex/docx-go/wml/body"
	"github.com/vortex/docx-go/wml/para"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
)

// ---------------------------------------------------------------------------
// Public types (exact match to C-31 contract)
// ---------------------------------------------------------------------------

// Severity indicates how critical a validation issue is.
type Severity int

const (
	Warn  Severity = iota // cosmetic or best-practice violation
	Error                 // Word will repair or display incorrectly
	Fatal                 // document cannot be opened at all
)

// String returns a human-readable label.
func (s Severity) String() string {
	switch s {
	case Warn:
		return "WARN"
	case Error:
		return "ERROR"
	case Fatal:
		return "FATAL"
	default:
		return "UNKNOWN"
	}
}

// Issue describes a single validation finding.
type Issue struct {
	Severity Severity
	Part     string // e.g. "word/document.xml"
	Path     string // e.g. "body/p[3]/r[1]/t"
	Code     string // e.g. "EMPTY_TC", "MISSING_REL"
	Message  string
}

// String formats the issue for logging / display.
func (i Issue) String() string {
	loc := i.Part
	if i.Path != "" {
		loc += "/" + i.Path
	}
	return fmt.Sprintf("[%s] %s (%s): %s", i.Severity, i.Code, loc, i.Message)
}

// ---------------------------------------------------------------------------
// Issue codes
// ---------------------------------------------------------------------------

const (
	CodeMissingDocumentBody = "MISSING_DOCUMENT_BODY"
	CodeMissingStyles       = "MISSING_STYLES"
	CodeMissingSettings     = "MISSING_SETTINGS"
	CodeMissingFonts        = "MISSING_FONTS"
	CodeMissingTheme        = "MISSING_THEME"
	CodeMissingWebSettings  = "MISSING_WEBSETTINGS"
	CodeEmptyBody           = "EMPTY_BODY"
	CodeEmptyTC             = "EMPTY_TC"
	CodeNilDocument         = "NIL_DOCUMENT"
)

// ---------------------------------------------------------------------------
// Validate inspects the document and returns all found issues.
// Reference: appendix 5.14 "full checklist for error-free opening"
// ---------------------------------------------------------------------------

// Validate checks a packaging.Document for structural problems.
// It returns a slice of Issue sorted roughly by severity (Fatal first).
func Validate(doc *packaging.Document) []Issue {
	if doc == nil {
		return []Issue{{
			Severity: Fatal,
			Part:     "",
			Code:     CodeNilDocument,
			Message:  "document pointer is nil",
		}}
	}

	var issues []Issue

	// --- required parts ------------------------------------------------
	issues = append(issues, checkRequiredParts(doc)...)

	// --- body-level checks ---------------------------------------------
	issues = append(issues, checkBody(doc)...)

	// sort: Fatal → Error → Warn for consumer convenience
	sortIssues(issues)

	return issues
}

// ---------------------------------------------------------------------------
// Rule: required parts (reference-appendix 5.14 – Parts section)
// ---------------------------------------------------------------------------

func checkRequiredParts(doc *packaging.Document) []Issue {
	var issues []Issue

	// Document + Body — without these the file is meaningless
	if doc.Document == nil {
		issues = append(issues, Issue{
			Severity: Fatal,
			Part:     "word/document.xml",
			Code:     CodeMissingDocumentBody,
			Message:  "CT_Document is nil — document part is missing or unparseable",
		})
	} else if doc.Document.Body == nil {
		issues = append(issues, Issue{
			Severity: Fatal,
			Part:     "word/document.xml",
			Code:     CodeMissingDocumentBody,
			Message:  "CT_Body is nil — <w:body> element is missing",
		})
	}

	// Styles — Word will regenerate, but output may look wrong
	if doc.Styles == nil {
		issues = append(issues, Issue{
			Severity: Error,
			Part:     "word/styles.xml",
			Code:     CodeMissingStyles,
			Message:  "styles part is missing — Word will recreate with defaults",
		})
	}

	// Settings — clrSchemeMapping etc. cause repair dialog if absent
	if doc.Settings == nil {
		issues = append(issues, Issue{
			Severity: Error,
			Part:     "word/settings.xml",
			Code:     CodeMissingSettings,
			Message:  "settings part is missing — Word may show repair dialog",
		})
	}

	// FontTable
	if doc.Fonts == nil {
		issues = append(issues, Issue{
			Severity: Error,
			Part:     "word/fontTable.xml",
			Code:     CodeMissingFonts,
			Message:  "fontTable part is missing — Word will recreate",
		})
	}

	// Theme — required when styles reference themeColor / asciiTheme
	if len(doc.Theme) == 0 {
		issues = append(issues, Issue{
			Severity: Error,
			Part:     "word/theme/theme1.xml",
			Code:     CodeMissingTheme,
			Message:  "theme part is missing or empty — styles referencing theme will break",
		})
	}

	// WebSettings — must exist even if empty
	if len(doc.WebSettings) == 0 {
		issues = append(issues, Issue{
			Severity: Warn,
			Part:     "word/webSettings.xml",
			Code:     CodeMissingWebSettings,
			Message:  "webSettings part is missing — Word adds it on repair",
		})
	}

	return issues
}

// ---------------------------------------------------------------------------
// Rule: body content — empty body, empty table cells
// ---------------------------------------------------------------------------

func checkBody(doc *packaging.Document) []Issue {
	if doc.Document == nil || doc.Document.Body == nil {
		return nil // already reported by checkRequiredParts
	}

	var issues []Issue
	bd := doc.Document.Body

	if len(bd.Content) == 0 {
		issues = append(issues, Issue{
			Severity: Warn,
			Part:     "word/document.xml",
			Path:     "body",
			Code:     CodeEmptyBody,
			Message:  "body has no block-level content",
		})
	}

	// Walk body content looking for tables with empty cells.
	issues = append(issues, walkBlockContent(bd.Content, "body")...)

	return issues
}

// walkBlockContent recursively inspects block-level elements.
func walkBlockContent(elems []shared.BlockLevelElement, basePath string) []Issue {
	var issues []Issue
	tblIdx := 0
	for _, el := range elems {
		switch v := el.(type) {
		case *table.CT_Tbl:
			tblPath := fmt.Sprintf("%s/tbl[%d]", basePath, tblIdx)
			issues = append(issues, checkTable(v, tblPath)...)
			tblIdx++
		}
	}
	return issues
}

// checkTable validates table structure — primarily the "each tc must have ≥ 1 p" rule.
func checkTable(tbl *table.CT_Tbl, tblPath string) []Issue {
	if tbl == nil {
		return nil
	}
	var issues []Issue

	ri := 0
	for _, tc := range tbl.Content {
		row, ok := tc.(table.CT_Row)
		if !ok {
			continue
		}
		rowPath := fmt.Sprintf("%s/tr[%d]", tblPath, ri)
		ri++

		ci := 0
		for _, rc := range row.Content {
			cell, ok := rc.(table.CT_Tc)
			if !ok {
				continue
			}
			cellPath := fmt.Sprintf("%s/tc[%d]", rowPath, ci)
			ci++

			if !cellHasParagraph(cell) {
				issues = append(issues, Issue{
					Severity: Error,
					Part:     "word/document.xml",
					Path:     cellPath,
					Code:     CodeEmptyTC,
					Message:  "table cell has no paragraph — Word requires at least one <w:p> per cell",
				})
			}

			// Recurse into nested tables inside the cell
			issues = append(issues, walkBlockContent(cell.Content, cellPath)...)
		}
	}
	return issues
}

// cellHasParagraph returns true if at least one BlockLevelElement in the cell is a paragraph.
func cellHasParagraph(tc table.CT_Tc) bool {
	for _, el := range tc.Content {
		if _, ok := el.(*para.CT_P); ok {
			return true
		}
	}
	return false
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// sortIssues places Fatal before Error before Warn (stable within group).
func sortIssues(issues []Issue) {
	// Simple stable in-place partition (small N expected).
	n := len(issues)
	for i := 1; i < n; i++ {
		for j := i; j > 0 && issues[j].Severity > issues[j-1].Severity; j-- {
			issues[j], issues[j-1] = issues[j-1], issues[j]
		}
	}
}

// HasFatal returns true if any issue is Fatal.
func HasFatal(issues []Issue) bool {
	for _, i := range issues {
		if i.Severity == Fatal {
			return true
		}
	}
	return false
}

// HasErrors returns true if any issue is Error or Fatal.
func HasErrors(issues []Issue) bool {
	for _, i := range issues {
		if i.Severity >= Error {
			return true
		}
	}
	return false
}

// Filter returns issues matching the given severity.
func Filter(issues []Issue, sev Severity) []Issue {
	var out []Issue
	for _, i := range issues {
		if i.Severity == sev {
			out = append(out, i)
		}
	}
	return out
}

// Summary returns a one-line human-readable summary like "2 errors, 1 warning".
func Summary(issues []Issue) string {
	var fatal, errs, warns int
	for _, i := range issues {
		switch i.Severity {
		case Fatal:
			fatal++
		case Error:
			errs++
		case Warn:
			warns++
		}
	}
	if fatal+errs+warns == 0 {
		return "document is valid"
	}
	var parts []string
	if fatal > 0 {
		parts = append(parts, fmt.Sprintf("%d fatal", fatal))
	}
	if errs > 0 {
		parts = append(parts, fmt.Sprintf("%d error(s)", errs))
	}
	if warns > 0 {
		parts = append(parts, fmt.Sprintf("%d warning(s)", warns))
	}
	return strings.Join(parts, ", ")
}

// Ensure body and table imports are used (they are used in type assertions above,
// but the compiler might complain if the switch cases are the only reference).
var (
	_ *body.CT_Document
	_ *table.CT_Tbl
	_ *para.CT_P
)
