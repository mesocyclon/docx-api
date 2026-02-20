// Package shared defines content-level interfaces and the RawXML type used
// across all WML sub-packages. It breaks cyclic dependencies: body, table,
// para, run, etc. all import shared but never import each other directly.
//
// See contracts.md C-05 and patterns.md §1 for the architectural rationale.
package shared

// BlockLevelElement represents a block-level item that may appear inside
// CT_Body, CT_Tc, CT_HdrFtr, CT_Comment, or CT_FtnEdn.
// Concrete implementations: CT_P (paragraph), CT_Tbl (table), and RawXML
// (unknown / extension elements preserved for round-trip fidelity).
type BlockLevelElement interface {
	blockLevelElement()
}

// ParagraphContent represents an inline item inside a paragraph (CT_P).
// Concrete implementations: CT_R (run), CT_Hyperlink, CT_BookmarkStart/End,
// CT_RunTrackChange (ins/del), and RawXML.
type ParagraphContent interface {
	paragraphContent()
}

// RunContent represents content inside a single run (CT_R).
// Concrete implementations: CT_Text, CT_Br, CT_Drawing, CT_FldChar, CT_Tab,
// and RawXML.
type RunContent interface {
	runContent()
}

// ---------------------------------------------------------------------------
// Embeddable markers — external packages embed these zero-sized structs to
// satisfy the unexported-method interfaces above.  Go requires unexported
// interface methods to originate in the defining package, so types in table,
// para, run, etc. embed the corresponding marker instead of declaring the
// method directly.
//
// Example usage in wml/table:
//
//	type CT_Tbl struct {
//	    shared.BlockLevelMarker   // ← implements shared.BlockLevelElement
//	    TblPr   *CT_TblPr
//	    ...
//	}
// ---------------------------------------------------------------------------

// BlockLevelMarker is embedded by types that implement BlockLevelElement.
type BlockLevelMarker struct{}

func (BlockLevelMarker) blockLevelElement() {}

// ParagraphContentMarker is embedded by types that implement ParagraphContent.
type ParagraphContentMarker struct{}

func (ParagraphContentMarker) paragraphContent() {}

// RunContentMarker is embedded by types that implement RunContent.
type RunContentMarker struct{}

func (RunContentMarker) runContent() {}
