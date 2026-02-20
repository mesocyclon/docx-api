package body

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/para"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
	"github.com/vortex/docx-go/xmltypes"
)

func init() {
	shared.RegisterBlockFactory(blockFactory)
}

// blockFactory creates typed block-level elements from XML names.
// It handles elements that the body package "owns": p, tbl, sdt.
func blockFactory(name xml.Name) shared.BlockLevelElement {
	if name.Space != xmltypes.NSw && name.Space != "" {
		return nil
	}
	switch name.Local {
	case "p":
		return ParagraphElement{P: &para.CT_P{}}
	case "tbl":
		return TableElement{T: &table.CT_Tbl{}}
	case "sdt":
		return SdtBlockElement{Sdt: &CT_SdtBlock{}}
	}
	return nil
}
