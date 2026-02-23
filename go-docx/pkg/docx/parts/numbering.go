package parts

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// NumberingPart is the proxy for the numbering.xml part containing numbering
// definitions for a document or glossary.
//
// Mirrors Python NumberingPart(XmlPart).
type NumberingPart struct {
	*opc.XmlPart
}

// NewNumberingPart wraps an XmlPart as a NumberingPart.
func NewNumberingPart(xp *opc.XmlPart) *NumberingPart {
	return &NumberingPart{XmlPart: xp}
}

// LoadNumberingPart is a PartConstructor for loading NumberingPart from a package.
func LoadNumberingPart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	xp, err := opc.NewXmlPart(partName, contentType, blob, pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: loading numbering part %q: %w", partName, err)
	}
	return NewNumberingPart(xp), nil
}
