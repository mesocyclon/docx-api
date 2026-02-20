package validator

import (
	"fmt"

	"github.com/vortex/docx-go/packaging"
	"github.com/vortex/docx-go/wml/para"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
)

// minimalWebSettings is the smallest valid webSettings part that Word accepts.
var minimalWebSettings = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
	`<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
	` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
	` mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh"/>`)

// AutoFix attempts to repair known structural issues in the document in-place.
// It returns a human-readable description of each fix applied.
// AutoFix is idempotent: calling it twice produces no additional fixes.
func AutoFix(doc *packaging.Document) []string {
	if doc == nil {
		return nil
	}

	var fixes []string

	fixes = append(fixes, fixWebSettings(doc)...)
	fixes = append(fixes, fixEmptyTableCells(doc)...)

	return fixes
}

// ---------------------------------------------------------------------------
// Fix: MISSING_WEBSETTINGS
// ---------------------------------------------------------------------------

func fixWebSettings(doc *packaging.Document) []string {
	if len(doc.WebSettings) > 0 {
		return nil
	}
	doc.WebSettings = minimalWebSettings
	return []string{"added minimal webSettings.xml part"}
}

// ---------------------------------------------------------------------------
// Fix: EMPTY_TC — insert an empty <w:p/> into every table cell that has none.
// Reference: appendix 5.14 "each tc → at least one p"
// ---------------------------------------------------------------------------

func fixEmptyTableCells(doc *packaging.Document) []string {
	if doc.Document == nil || doc.Document.Body == nil {
		return nil
	}
	count := fixBlockContent(doc.Document.Body.Content)
	if count == 0 {
		return nil
	}
	return []string{fmt.Sprintf("added empty paragraph to %d empty table cell(s)", count)}
}

// fixBlockContent walks block-level elements and fixes empty cells.
// Returns the number of cells fixed.
func fixBlockContent(elems []shared.BlockLevelElement) int {
	fixed := 0
	for _, el := range elems {
		tbl, ok := el.(*table.CT_Tbl)
		if !ok || tbl == nil {
			continue
		}
		for ti, tc := range tbl.Content {
			row, ok := tc.(table.CT_Row)
			if !ok {
				continue
			}
			for ci, rc := range row.Content {
				cell, ok := rc.(table.CT_Tc)
				if !ok {
					continue
				}

				if !cellHasParagraph(cell) {
					// Prepend an empty paragraph so that Word is happy.
					emptyP := &para.CT_P{}
					cell.Content = append([]shared.BlockLevelElement{emptyP}, cell.Content...)
					// Write the modified cell back (CT_Tc is a value type in the slice).
					row.Content[ci] = cell
					fixed++
				}

				// Recurse into nested tables inside the cell.
				fixed += fixBlockContent(cell.Content)
			}
			// Write the modified row back (CT_Row is a value type in the slice).
			tbl.Content[ti] = row
		}
	}
	return fixed
}
