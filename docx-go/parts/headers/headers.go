// Package headers provides Parse / Serialize functions for OOXML header and
// footer parts (word/header*.xml, word/footer*.xml).
//
// It is a thin part-level wrapper around wml/hdft.CT_HdrFtr.  Both header
// and footer parts share the same schema (CT_HdrFtr); the only difference
// is the root element name ("w:hdr" vs "w:ftr") and the content type.
//
// See contracts.md C-27 and reference-appendix §3.7.
package headers

import (
	"github.com/vortex/docx-go/wml/hdft"
)

// ---------------------------------------------------------------------------
// Parse
// ---------------------------------------------------------------------------

// Parse deserialises raw XML bytes from a header or footer part into a
// CT_HdrFtr.  The root element may be either <w:hdr> or <w:ftr>.
func Parse(data []byte) (*hdft.CT_HdrFtr, error) {
	return hdft.Parse(data)
}

// ---------------------------------------------------------------------------
// Serialize — Header
// ---------------------------------------------------------------------------

// Serialize marshals CT_HdrFtr as a header part (root element <w:hdr>).
func Serialize(hf *hdft.CT_HdrFtr) ([]byte, error) {
	return hdft.Serialize(hf, "w:hdr")
}

// ---------------------------------------------------------------------------
// Serialize — Footer
// ---------------------------------------------------------------------------

// SerializeFooter marshals CT_HdrFtr as a footer part (root element <w:ftr>).
func SerializeFooter(hf *hdft.CT_HdrFtr) ([]byte, error) {
	return hdft.Serialize(hf, "w:ftr")
}
