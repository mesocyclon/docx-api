// Package para implements the CT_P (paragraph) element and its direct children
// for the OOXML WordprocessingML schema.
//
// A paragraph contains an optional pPr (paragraph properties) element followed
// by a mixed sequence of inline content: runs, hyperlinks, bookmarks, tracked
// changes, structured document tags, and unknown extensions.
//
// See contracts.md C-16 and reference-appendix.md §3.1.
package para

import (
	"encoding/xml"
	"strconv"

	"github.com/vortex/docx-go/wml/ppr"
	"github.com/vortex/docx-go/wml/run"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/tracking"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_P represents a paragraph (<w:p>) in a WordprocessingML document.
// It implements shared.BlockLevelElement so it can appear inside CT_Body,
// CT_Tc, CT_HdrFtr, etc.
type CT_P struct {
	shared.BlockLevelMarker

	PPr     *ppr.CT_PPr
	Content []shared.ParagraphContent

	// Attributes
	RsidR        *string // w:rsidR
	RsidRDefault *string // w:rsidRDefault
	RsidP        *string // w:rsidP
	RsidRPr      *string // w:rsidRPr
	RsidDel      *string // w:rsidDel
	ParaId       *string // w14:paraId
	TextId       *string // w14:textId
}

// ---------------------------------------------------------------------------
// MarshalXML — CT_P
// ---------------------------------------------------------------------------

func (p *CT_P) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Build start element with correct namespace and attributes.
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "p"}
	appendStringAttr(&start, xmltypes.NSw, "rsidR", p.RsidR)
	appendStringAttr(&start, xmltypes.NSw, "rsidRDefault", p.RsidRDefault)
	appendStringAttr(&start, xmltypes.NSw, "rsidP", p.RsidP)
	appendStringAttr(&start, xmltypes.NSw, "rsidRPr", p.RsidRPr)
	appendStringAttr(&start, xmltypes.NSw, "rsidDel", p.RsidDel)
	appendStringAttr(&start, xmltypes.NSw14, "paraId", p.ParaId)
	appendStringAttr(&start, xmltypes.NSw14, "textId", p.TextId)

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 1. pPr — always first child if present.
	if p.PPr != nil {
		if err := e.EncodeElement(p.PPr, xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "pPr"},
		}); err != nil {
			return err
		}
	}

	// 2. Content items — in document order.
	if err := marshalParagraphContent(e, p.Content); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// ---------------------------------------------------------------------------
// UnmarshalXML — CT_P
// ---------------------------------------------------------------------------

func (p *CT_P) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes.
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidR":
			s := attr.Value
			p.RsidR = &s
		case "rsidRDefault":
			s := attr.Value
			p.RsidRDefault = &s
		case "rsidP":
			s := attr.Value
			p.RsidP = &s
		case "rsidRPr":
			s := attr.Value
			p.RsidRPr = &s
		case "rsidDel":
			s := attr.Value
			p.RsidDel = &s
		case "paraId":
			s := attr.Value
			p.ParaId = &s
		case "textId":
			s := attr.Value
			p.TextId = &s
		}
	}

	// Parse children.
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "pPr" && isWMLNamespace(t.Name.Space) {
				p.PPr = &ppr.CT_PPr{}
				if err := d.DecodeElement(p.PPr, &t); err != nil {
					return err
				}
				continue
			}

			// Delegate to the paragraph-content dispatcher.
			item := unmarshalOneParagraphContent(d, t)
			if item != nil {
				p.Content = append(p.Content, item)
			}

		case xml.EndElement:
			return nil
		}
	}
}

// ---------------------------------------------------------------------------
// marshalParagraphContent — encode a slice of ParagraphContent items.
// Used by CT_P, CT_Hyperlink, CT_SimpleField, and CT_SdtRun.
// ---------------------------------------------------------------------------

func marshalParagraphContent(e *xml.Encoder, items []shared.ParagraphContent) error {
	for _, item := range items {
		switch v := item.(type) {
		case RunItem:
			if v.R != nil {
				if err := e.EncodeElement(v.R, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "r"},
				}); err != nil {
					return err
				}
			}
		case *RunItem:
			if v != nil && v.R != nil {
				if err := e.EncodeElement(v.R, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "r"},
				}); err != nil {
					return err
				}
			}

		case HyperlinkItem:
			if v.H != nil {
				if err := e.EncodeElement(v.H, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "hyperlink"},
				}); err != nil {
					return err
				}
			}
		case *HyperlinkItem:
			if v != nil && v.H != nil {
				if err := e.EncodeElement(v.H, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "hyperlink"},
				}); err != nil {
					return err
				}
			}

		case SimpleFieldItem:
			if v.F != nil {
				if err := e.EncodeElement(v.F, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "fldSimple"},
				}); err != nil {
					return err
				}
			}
		case *SimpleFieldItem:
			if v != nil && v.F != nil {
				if err := e.EncodeElement(v.F, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "fldSimple"},
				}); err != nil {
					return err
				}
			}

		case InsItem:
			if v.Ins != nil {
				if err := marshalTrackChange(e, "ins", v.Ins); err != nil {
					return err
				}
			}
		case *InsItem:
			if v != nil && v.Ins != nil {
				if err := marshalTrackChange(e, "ins", v.Ins); err != nil {
					return err
				}
			}

		case DelItem:
			if v.Del != nil {
				if err := marshalTrackChange(e, "del", v.Del); err != nil {
					return err
				}
			}
		case *DelItem:
			if v != nil && v.Del != nil {
				if err := marshalTrackChange(e, "del", v.Del); err != nil {
					return err
				}
			}

		case BookmarkStartItem:
			if v.B != nil {
				if err := e.EncodeElement(v.B, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "bookmarkStart"},
				}); err != nil {
					return err
				}
			}
		case *BookmarkStartItem:
			if v != nil && v.B != nil {
				if err := e.EncodeElement(v.B, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "bookmarkStart"},
				}); err != nil {
					return err
				}
			}

		case BookmarkEndItem:
			if v.B != nil {
				if err := e.EncodeElement(v.B, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "bookmarkEnd"},
				}); err != nil {
					return err
				}
			}
		case *BookmarkEndItem:
			if v != nil && v.B != nil {
				if err := e.EncodeElement(v.B, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "bookmarkEnd"},
				}); err != nil {
					return err
				}
			}

		case CommentRangeStartItem:
			if v.C != nil {
				if err := e.EncodeElement(v.C, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "commentRangeStart"},
				}); err != nil {
					return err
				}
			}
		case *CommentRangeStartItem:
			if v != nil && v.C != nil {
				if err := e.EncodeElement(v.C, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "commentRangeStart"},
				}); err != nil {
					return err
				}
			}

		case CommentRangeEndItem:
			if v.C != nil {
				if err := e.EncodeElement(v.C, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "commentRangeEnd"},
				}); err != nil {
					return err
				}
			}
		case *CommentRangeEndItem:
			if v != nil && v.C != nil {
				if err := e.EncodeElement(v.C, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "commentRangeEnd"},
				}); err != nil {
					return err
				}
			}

		case SdtRunItem:
			if v.Sdt != nil {
				if err := e.EncodeElement(v.Sdt, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "sdt"},
				}); err != nil {
					return err
				}
			}
		case *SdtRunItem:
			if v != nil && v.Sdt != nil {
				if err := e.EncodeElement(v.Sdt, xml.StartElement{
					Name: xml.Name{Space: xmltypes.NSw, Local: "sdt"},
				}); err != nil {
					return err
				}
			}

		case RawParagraphContent:
			if err := e.EncodeElement(v.Raw, xml.StartElement{Name: v.Raw.XMLName}); err != nil {
				return err
			}
		case *RawParagraphContent:
			if v != nil {
				if err := e.EncodeElement(v.Raw, xml.StartElement{Name: v.Raw.XMLName}); err != nil {
					return err
				}
			}

		case shared.RawXML:
			if err := e.EncodeElement(v, xml.StartElement{Name: v.XMLName}); err != nil {
				return err
			}
		}
	}
	return nil
}

// marshalTrackChange encodes a CT_RunTrackChange as <w:ins> or <w:del>.
// Note: CT_RunTrackChange.Content is []interface{} per contract C-14,
// so we convert elements individually to ParagraphContent for marshalling.
func marshalTrackChange(e *xml.Encoder, localName string, tc *tracking.CT_RunTrackChange) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: localName},
	}
	start.Attr = append(start.Attr,
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "id"}, Value: strconv.Itoa(tc.ID)},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "author"}, Value: tc.Author},
	)
	if tc.Date != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "date"}, Value: *tc.Date},
		)
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// CT_RunTrackChange.Content is []interface{} — may contain
	// ParagraphContent items (runs, etc.) or shared.RawXML.
	for _, item := range tc.Content {
		if pc, ok := item.(shared.ParagraphContent); ok {
			// shared.RawXML also satisfies ParagraphContent, so this
			// single branch handles both typed items and raw round-trip.
			if err := marshalParagraphContent(e, []shared.ParagraphContent{pc}); err != nil {
				return err
			}
		} else {
			// Best-effort for anything else stored in []interface{}.
			if err := e.EncodeElement(item, xml.StartElement{
				Name: xml.Name{Space: xmltypes.NSw, Local: "r"},
			}); err != nil {
				return err
			}
		}
	}

	return e.EncodeToken(start.End())
}

// ---------------------------------------------------------------------------
// unmarshalParagraphContent — decode children until EndElement.
// Used by CT_Hyperlink, CT_SimpleField, CT_SdtRun's sdtContent.
// ---------------------------------------------------------------------------

func unmarshalParagraphContent(d *xml.Decoder) ([]shared.ParagraphContent, error) {
	var items []shared.ParagraphContent
	for {
		tok, err := d.Token()
		if err != nil {
			return items, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			item := unmarshalOneParagraphContent(d, t)
			if item != nil {
				items = append(items, item)
			}
		case xml.EndElement:
			return items, nil
		}
	}
}

// unmarshalOneParagraphContent decodes a single child element into the
// appropriate ParagraphContent type.  Returns nil only if decoding fails
// silently (which should not happen in practice).
func unmarshalOneParagraphContent(d *xml.Decoder, t xml.StartElement) shared.ParagraphContent {
	// Check WML namespace elements first.
	switch t.Name.Local {
	case "r":
		r := &run.CT_R{}
		if err := d.DecodeElement(r, &t); err != nil {
			return captureRaw(d, t)
		}
		return RunItem{R: r}

	case "hyperlink":
		h := &CT_Hyperlink{}
		if err := d.DecodeElement(h, &t); err != nil {
			return captureRaw(d, t)
		}
		return HyperlinkItem{H: h}

	case "fldSimple":
		f := &CT_SimpleField{}
		if err := d.DecodeElement(f, &t); err != nil {
			return captureRaw(d, t)
		}
		return SimpleFieldItem{F: f}

	case "ins":
		ins := &tracking.CT_RunTrackChange{}
		if err := d.DecodeElement(ins, &t); err != nil {
			return captureRaw(d, t)
		}
		return InsItem{Ins: ins}

	case "del":
		del := &tracking.CT_RunTrackChange{}
		if err := d.DecodeElement(del, &t); err != nil {
			return captureRaw(d, t)
		}
		return DelItem{Del: del}

	case "bookmarkStart":
		b := &tracking.CT_Bookmark{}
		if err := d.DecodeElement(b, &t); err != nil {
			return captureRaw(d, t)
		}
		return BookmarkStartItem{B: b}

	case "bookmarkEnd":
		b := &tracking.CT_MarkupRange{}
		if err := d.DecodeElement(b, &t); err != nil {
			return captureRaw(d, t)
		}
		return BookmarkEndItem{B: b}

	case "commentRangeStart":
		c := &tracking.CT_Markup{}
		if err := d.DecodeElement(c, &t); err != nil {
			return captureRaw(d, t)
		}
		return CommentRangeStartItem{C: c}

	case "commentRangeEnd":
		c := &tracking.CT_Markup{}
		if err := d.DecodeElement(c, &t); err != nil {
			return captureRaw(d, t)
		}
		return CommentRangeEndItem{C: c}

	case "moveFrom", "moveTo":
		// Track-change moves share the same schema as ins/del.
		// Preserve as raw to maintain the correct element name on round-trip.
		return captureRaw(d, t)

	case "sdt":
		sdt := &CT_SdtRun{}
		if err := d.DecodeElement(sdt, &t); err != nil {
			return captureRaw(d, t)
		}
		return SdtRunItem{Sdt: sdt}

	default:
		// Unknown element — preserve as RawXML for round-trip fidelity.
		return captureRaw(d, t)
	}
}

// captureRaw decodes an element into shared.RawXML and wraps it as
// RawParagraphContent.
func captureRaw(d *xml.Decoder, t xml.StartElement) RawParagraphContent {
	var raw shared.RawXML
	_ = d.DecodeElement(&raw, &t)
	return RawParagraphContent{Raw: raw}
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// appendStringAttr adds a string attribute to a StartElement if the pointer
// is non-nil.
func appendStringAttr(start *xml.StartElement, space, local string, val *string) {
	if val != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: space, Local: local},
			Value: *val,
		})
	}
}

// isWMLNamespace returns true if the namespace is the WML namespace (or
// empty, which encoding/xml sometimes uses when namespace is inherited).
func isWMLNamespace(ns string) bool {
	return ns == "" || ns == xmltypes.NSw
}
