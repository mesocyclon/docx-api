package para

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
)

func init() {
	// Register CT_P as a block-level element so that body, tc, hdrFtr, etc.
	// can create it during unmarshal.
	shared.RegisterBlockFactory(func(name xml.Name) shared.BlockLevelElement {
		if name.Local == "p" && isWMLNamespace(name.Space) {
			return &CT_P{}
		}
		return nil
	})

	// Register paragraph-content types so that containers holding
	// []shared.ParagraphContent (e.g. CT_RunTrackChange inside tracking)
	// can create them during unmarshal via shared.CreateParagraphContent.
	shared.RegisterParagraphContentFactory(func(name xml.Name) shared.ParagraphContent {
		if !isWMLNamespace(name.Space) {
			return nil
		}
		switch name.Local {
		case "r":
			return &RunItem{}
		case "hyperlink":
			return &HyperlinkItem{}
		case "fldSimple":
			return &SimpleFieldItem{}
		case "bookmarkStart":
			return &BookmarkStartItem{}
		case "bookmarkEnd":
			return &BookmarkEndItem{}
		case "commentRangeStart":
			return &CommentRangeStartItem{}
		case "commentRangeEnd":
			return &CommentRangeEndItem{}
		case "sdt":
			return &SdtRunItem{}
		}
		return nil
	})
}
