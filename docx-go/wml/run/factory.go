package run

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

func init() {
	// Register CT_R as a ParagraphContent factory so that the paragraph
	// module can create runs when it encounters <w:r> during unmarshal.
	shared.RegisterParagraphContentFactory(func(name xml.Name) shared.ParagraphContent {
		if name.Local == "r" && (name.Space == xmltypes.NSw || name.Space == "" || name.Space == "w") {
			return &CT_R{}
		}
		return nil
	})

	// Register a RunContent factory so that other modules can create typed
	// run-content elements by XML name.
	shared.RegisterRunContentFactory(runContentFactory)
}

// runContentFactory creates a typed RunContent from an XML element name.
// Returns nil for unrecognised names.
func runContentFactory(name xml.Name) shared.RunContent {
	switch name.Local {
	case "t", "delText":
		return &CT_Text{}
	case "br":
		return &CT_Br{}
	case "drawing":
		return &CT_Drawing{}
	case "fldChar":
		return &CT_FldChar{}
	case "instrText", "delInstrText":
		return &CT_InstrText{}
	case "sym":
		return &CT_Sym{}
	case "footnoteReference", "endnoteReference", "commentReference":
		return &CT_FtnEdnRef{XMLName: name}
	default:
		if emptyElementNames[name.Local] {
			return &CT_EmptyRunContent{XMLName: name}
		}
		return nil
	}
}
