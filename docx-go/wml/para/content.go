package para

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/run"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/tracking"
)

// ---------------------------------------------------------------------------
// ParagraphContent implementations
//
// Each wrapper type embeds shared.ParagraphContentMarker to satisfy the
// shared.ParagraphContent interface (which has an unexported method).
//
// Types returned by the factory (RegisterParagraphContentFactory) also
// implement xml.Unmarshaler so that callers can d.DecodeElement(el, &start).
// ---------------------------------------------------------------------------

// RunItem wraps a run element (<w:r>).
type RunItem struct {
	shared.ParagraphContentMarker
	R *run.CT_R
}

func (ri *RunItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	ri.R = &run.CT_R{}
	return d.DecodeElement(ri.R, &start)
}

// HyperlinkItem wraps a hyperlink element (<w:hyperlink>).
type HyperlinkItem struct {
	shared.ParagraphContentMarker
	H *CT_Hyperlink
}

func (hi *HyperlinkItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	hi.H = &CT_Hyperlink{}
	return d.DecodeElement(hi.H, &start)
}

// SimpleFieldItem wraps a simple field element (<w:fldSimple>).
type SimpleFieldItem struct {
	shared.ParagraphContentMarker
	F *CT_SimpleField
}

func (si *SimpleFieldItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	si.F = &CT_SimpleField{}
	return d.DecodeElement(si.F, &start)
}

// InsItem wraps an insertion track-change element (<w:ins>).
type InsItem struct {
	shared.ParagraphContentMarker
	Ins *tracking.CT_RunTrackChange
}

func (ii *InsItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	ii.Ins = &tracking.CT_RunTrackChange{}
	return d.DecodeElement(ii.Ins, &start)
}

// DelItem wraps a deletion track-change element (<w:del>).
type DelItem struct {
	shared.ParagraphContentMarker
	Del *tracking.CT_RunTrackChange
}

func (di *DelItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	di.Del = &tracking.CT_RunTrackChange{}
	return d.DecodeElement(di.Del, &start)
}

// BookmarkStartItem wraps a bookmarkStart element (<w:bookmarkStart>).
type BookmarkStartItem struct {
	shared.ParagraphContentMarker
	B *tracking.CT_Bookmark
}

func (bi *BookmarkStartItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	bi.B = &tracking.CT_Bookmark{}
	return d.DecodeElement(bi.B, &start)
}

// BookmarkEndItem wraps a bookmarkEnd element (<w:bookmarkEnd>).
type BookmarkEndItem struct {
	shared.ParagraphContentMarker
	B *tracking.CT_MarkupRange
}

func (bi *BookmarkEndItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	bi.B = &tracking.CT_MarkupRange{}
	return d.DecodeElement(bi.B, &start)
}

// CommentRangeStartItem wraps a commentRangeStart element.
type CommentRangeStartItem struct {
	shared.ParagraphContentMarker
	C *tracking.CT_Markup
}

func (ci *CommentRangeStartItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	ci.C = &tracking.CT_Markup{}
	return d.DecodeElement(ci.C, &start)
}

// CommentRangeEndItem wraps a commentRangeEnd element.
type CommentRangeEndItem struct {
	shared.ParagraphContentMarker
	C *tracking.CT_Markup
}

func (ci *CommentRangeEndItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	ci.C = &tracking.CT_Markup{}
	return d.DecodeElement(ci.C, &start)
}

// SdtRunItem wraps a structured document tag (<w:sdt>) at paragraph level.
type SdtRunItem struct {
	shared.ParagraphContentMarker
	Sdt *CT_SdtRun
}

func (si *SdtRunItem) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	si.Sdt = &CT_SdtRun{}
	return d.DecodeElement(si.Sdt, &start)
}

// RawParagraphContent preserves an unrecognised element for round-trip.
type RawParagraphContent struct {
	shared.ParagraphContentMarker
	Raw shared.RawXML
}

func (rc *RawParagraphContent) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	return d.DecodeElement(&rc.Raw, &start)
}
