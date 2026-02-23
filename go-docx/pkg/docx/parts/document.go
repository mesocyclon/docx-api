package parts

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// DocumentPart is the main document part of a WordprocessingML package.
// It acts as broker to other parts such as image, core properties, and
// style parts. It also acts as a convenient delegate when a mid-document
// object needs a service involving a remote ancestor.
//
// Mirrors Python DocumentPart(StoryPart).
type DocumentPart struct {
	StoryPart
	// numberingPart is the only lazyproperty-cached part in Python.
	// _styles_part, _settings_part, _comments_part are @property (not
	// lazyproperty) in Python — they re-check the relationship each call.
	// The relationship graph itself acts as the cache.
	numberingPart *NumberingPart
}

// NewDocumentPart creates a DocumentPart wrapping the given XmlPart.
func NewDocumentPart(xp *opc.XmlPart) *DocumentPart {
	dp := &DocumentPart{
		StoryPart: StoryPart{XmlPart: xp},
	}
	// The document part is its own document part.
	dp.StoryPart.SetDocumentPart(dp)
	return dp
}

// --------------------------------------------------------------------------
// Body access
// --------------------------------------------------------------------------

// Body returns the CT_Body element of this document.
func (dp *DocumentPart) Body() (*oxml.CT_Body, error) {
	el := dp.Element()
	if el == nil {
		return nil, fmt.Errorf("parts: document element is nil")
	}
	doc := &oxml.CT_Document{Element: oxml.Element{E: el}}
	body := doc.Body()
	if body == nil {
		return nil, fmt.Errorf("parts: document has no body element")
	}
	return body, nil
}

// --------------------------------------------------------------------------
// Header / Footer
// --------------------------------------------------------------------------

// AddHeaderPart creates a new header part, relates it to this document part,
// and returns the header part and its relationship ID.
//
// Mirrors Python DocumentPart.add_header_part:
//
//	header_part = HeaderPart.new(self.package)
//	rId = self.relate_to(header_part, RT.HEADER)
//	return header_part, rId
func (dp *DocumentPart) AddHeaderPart() (*HeaderPart, string, error) {
	pkg := dp.Package()
	if pkg == nil {
		return nil, "", fmt.Errorf("parts: document part has no package")
	}
	hp, err := NewHeaderPart(pkg)
	if err != nil {
		return nil, "", fmt.Errorf("parts: creating header part: %w", err)
	}
	// Go-specific: add to parts map so NextPartname sees it for subsequent calls.
	// Python discovers parts via relationship graph traversal in next_partname.
	pkg.AddPart(hp)
	rel := dp.Rels().GetOrAdd(opc.RTHeader, hp)
	return hp, rel.RID, nil
}

// AddFooterPart creates a new footer part, relates it to this document part,
// and returns the footer part and its relationship ID.
//
// Mirrors Python DocumentPart.add_footer_part.
func (dp *DocumentPart) AddFooterPart() (*FooterPart, string, error) {
	pkg := dp.Package()
	if pkg == nil {
		return nil, "", fmt.Errorf("parts: document part has no package")
	}
	fp, err := NewFooterPart(pkg)
	if err != nil {
		return nil, "", fmt.Errorf("parts: creating footer part: %w", err)
	}
	pkg.AddPart(fp)
	rel := dp.Rels().GetOrAdd(opc.RTFooter, fp)
	return fp, rel.RID, nil
}

// DropHeaderPart removes the header part relationship identified by rId.
// Uses reference-count-aware deletion matching Python DocumentPart.drop_header_part.
//
// Mirrors Python: self.drop_rel(rId)
func (dp *DocumentPart) DropHeaderPart(rId string) {
	dp.DropRel(rId)
}

// HeaderPartByRID returns the HeaderPart related by the given rId.
//
// Mirrors Python DocumentPart.header_part(rId) → self.related_parts[rId].
func (dp *DocumentPart) HeaderPartByRID(rId string) (*HeaderPart, error) {
	rel := dp.Rels().GetByRID(rId)
	if rel == nil {
		return nil, fmt.Errorf("parts: no relationship %q", rId)
	}
	if rel.TargetPart == nil {
		return nil, fmt.Errorf("parts: relationship %q has no target part", rId)
	}
	hp, ok := rel.TargetPart.(*HeaderPart)
	if !ok {
		return nil, fmt.Errorf("parts: relationship %q target is %T, want *HeaderPart", rId, rel.TargetPart)
	}
	return hp, nil
}

// FooterPartByRID returns the FooterPart related by the given rId.
//
// Mirrors Python DocumentPart.footer_part(rId) → self.related_parts[rId].
func (dp *DocumentPart) FooterPartByRID(rId string) (*FooterPart, error) {
	rel := dp.Rels().GetByRID(rId)
	if rel == nil {
		return nil, fmt.Errorf("parts: no relationship %q", rId)
	}
	if rel.TargetPart == nil {
		return nil, fmt.Errorf("parts: relationship %q has no target part", rId)
	}
	fp, ok := rel.TargetPart.(*FooterPart)
	if !ok {
		return nil, fmt.Errorf("parts: relationship %q target is %T, want *FooterPart", rId, rel.TargetPart)
	}
	return fp, nil
}

// --------------------------------------------------------------------------
// StylesPart — @property in Python (NOT lazyproperty), re-checks each call
// --------------------------------------------------------------------------

// StylesPart returns the StylesPart for this document, creating a default
// one if not present. NOT cached in a struct field — the relationship graph
// acts as the cache, matching Python's @property behavior.
//
// Mirrors Python DocumentPart._styles_part property.
func (dp *DocumentPart) StylesPart() (*StylesPart, error) {
	rel, err := dp.Rels().GetByRelType(opc.RTStyles)
	if err == nil && rel.TargetPart != nil {
		if sp, ok := rel.TargetPart.(*StylesPart); ok {
			return sp, nil
		}
	}
	// Not found — create default, relate, return.
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	sp, err := DefaultStylesPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default styles part: %w", err)
	}
	pkg.AddPart(sp)
	dp.Rels().GetOrAdd(opc.RTStyles, sp)
	return sp, nil
}

// Styles returns the CT_Styles element from the styles part.
//
// Mirrors Python DocumentPart.styles → self._styles_part.styles.
func (dp *DocumentPart) Styles() (*oxml.CT_Styles, error) {
	sp, err := dp.StylesPart()
	if err != nil {
		return nil, err
	}
	return sp.Styles()
}

// --------------------------------------------------------------------------
// NumberingPart — @lazyproperty in Python (the ONLY cached one)
// --------------------------------------------------------------------------

// NumberingPart returns the NumberingPart for this document. Unlike styles
// and settings, numbering does not auto-create a default part in the Python
// source (NumberingPart.new() raises NotImplementedError). We only resolve
// existing relationships. Cached per Python lazyproperty.
//
// Mirrors Python DocumentPart.numbering_part (lazyproperty).
func (dp *DocumentPart) NumberingPart() (*NumberingPart, error) {
	if dp.numberingPart != nil {
		return dp.numberingPart, nil
	}
	rel, err := dp.Rels().GetByRelType(opc.RTNumbering)
	if err != nil {
		return nil, fmt.Errorf("parts: no numbering part: %w", err)
	}
	if rel.TargetPart == nil {
		return nil, fmt.Errorf("parts: numbering relationship has no target part")
	}
	np, ok := rel.TargetPart.(*NumberingPart)
	if !ok {
		return nil, fmt.Errorf("parts: numbering target is %T, want *NumberingPart", rel.TargetPart)
	}
	dp.numberingPart = np
	return np, nil
}

// --------------------------------------------------------------------------
// SettingsPart — @property in Python (NOT lazyproperty)
// --------------------------------------------------------------------------

// SettingsPart returns the SettingsPart for this document, creating a
// default one if not present.
//
// Mirrors Python DocumentPart._settings_part property.
func (dp *DocumentPart) SettingsPart() (*SettingsPart, error) {
	rel, err := dp.Rels().GetByRelType(opc.RTSettings)
	if err == nil && rel.TargetPart != nil {
		if sp, ok := rel.TargetPart.(*SettingsPart); ok {
			return sp, nil
		}
	}
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	sp, err := DefaultSettingsPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default settings part: %w", err)
	}
	pkg.AddPart(sp)
	dp.Rels().GetOrAdd(opc.RTSettings, sp)
	return sp, nil
}

// Settings returns the CT_Settings element from the settings part.
//
// Mirrors Python DocumentPart.settings → self._settings_part.settings.
func (dp *DocumentPart) Settings() (*oxml.CT_Settings, error) {
	sp, err := dp.SettingsPart()
	if err != nil {
		return nil, err
	}
	return sp.SettingsElement()
}

// --------------------------------------------------------------------------
// CommentsPart — @property in Python (NOT lazyproperty)
// --------------------------------------------------------------------------

// CommentsPart returns the CommentsPart for this document, creating a
// default one if not present.
//
// Mirrors Python DocumentPart._comments_part property.
func (dp *DocumentPart) CommentsPart() (*CommentsPart, error) {
	rel, err := dp.Rels().GetByRelType(opc.RTComments)
	if err == nil && rel.TargetPart != nil {
		if cp, ok := rel.TargetPart.(*CommentsPart); ok {
			return cp, nil
		}
	}
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	cp, err := DefaultCommentsPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("parts: creating default comments part: %w", err)
	}
	pkg.AddPart(cp)
	dp.Rels().GetOrAdd(opc.RTComments, cp)
	return cp, nil
}

// CommentsElement returns the CT_Comments element from the comments part.
//
// Mirrors Python DocumentPart.comments (element access portion — the domain
// Comments object is MR-11).
func (dp *DocumentPart) CommentsElement() (*oxml.CT_Comments, error) {
	cp, err := dp.CommentsPart()
	if err != nil {
		return nil, err
	}
	return cp.CommentsElement()
}

// --------------------------------------------------------------------------
// CoreProperties
// --------------------------------------------------------------------------

// CoreProperties returns the part related by RTCoreProperties.
//
// Mirrors Python DocumentPart.core_properties → self.package.core_properties.
func (dp *DocumentPart) CoreProperties() (opc.Part, error) {
	pkg := dp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: document part has no package")
	}
	return pkg.RelatedPart(opc.RTCoreProperties)
}

// --------------------------------------------------------------------------
// Style delegation
// --------------------------------------------------------------------------

// GetStyle returns the style matching styleID and styleType.
// If styleID is nil, the default style for styleType is returned.
// If styleID does not match a defined style of styleType, the default
// style for styleType is returned.
//
// Mirrors Python DocumentPart.get_style → self.styles.get_by_id(style_id, style_type).
//
// Python Styles._get_by_id:
//
//	style = self._element.get_by_id(style_id)
//	if style is None or style.type != style_type:
//	    return self.default(style_type)
func (dp *DocumentPart) GetStyle(styleID *string, styleType enum.WdStyleType) (*oxml.CT_Style, error) {
	ss, err := dp.Styles()
	if err != nil {
		return nil, err
	}
	if styleID == nil {
		return ss.DefaultFor(styleType), nil
	}
	s := ss.GetByID(*styleID)
	xmlType, _ := styleType.ToXml()
	if s == nil || s.Type() != xmlType {
		// Fall back to default for the style type (matches Python _get_by_id).
		return ss.DefaultFor(styleType), nil
	}
	return s, nil
}

// GetStyleID returns the style_id string for styleOrName of styleType.
// Returns nil if styleOrName is nil or if the resolved style is the default
// for styleType.
//
// Mirrors Python DocumentPart.get_style_id → self.styles.get_style_id(style_or_name, style_type).
//
// Python _get_style_id_from_style:
//
//	if style.type != style_type:
//	    raise ValueError("assigned style is type %s, need type %s")
//	if style == self.default(style_type):
//	    return None
//	return style.style_id
//
// NOTE: Full BaseStyle support (accepting style objects) is added in MR-11.
// For MR-9 this handles string name lookups.
func (dp *DocumentPart) GetStyleID(styleOrName interface{}, styleType enum.WdStyleType) (*string, error) {
	if styleOrName == nil {
		return nil, nil
	}
	ss, err := dp.Styles()
	if err != nil {
		return nil, err
	}
	name, ok := styleOrName.(string)
	if !ok {
		return nil, fmt.Errorf("parts: GetStyleID expects string or nil (BaseStyle support in MR-11), got %T", styleOrName)
	}
	s := ss.GetByName(name)
	if s == nil {
		return nil, fmt.Errorf("parts: no style with name %q", name)
	}
	// Check style type matches (Python raises ValueError on mismatch).
	xmlType, _ := styleType.ToXml()
	if s.Type() != xmlType {
		return nil, fmt.Errorf("parts: style %q is type %q, need type %q", name, s.Type(), xmlType)
	}
	// If this style is the default for its type, return nil (matches Python).
	def := ss.DefaultFor(styleType)
	if def != nil && def.E == s.E {
		return nil, nil
	}
	id := s.StyleId()
	return &id, nil
}

// --------------------------------------------------------------------------
// InlineShapes (element access only — domain object is MR-11)
// --------------------------------------------------------------------------

// InlineShapeElements returns all wp:inline elements found within the
// document body. This provides the raw element access; the domain
// InlineShapes proxy is created in MR-11.
//
// Mirrors the element query underlying Python DocumentPart.inline_shapes.
func (dp *DocumentPart) InlineShapeElements() ([]*etree.Element, error) {
	body, err := dp.Body()
	if err != nil {
		return nil, err
	}
	var inlines []*etree.Element
	findInlines(body.E, &inlines)
	return inlines, nil
}

// findInlines recursively collects wp:inline elements.
func findInlines(el *etree.Element, result *[]*etree.Element) {
	if el.Tag == "inline" && (el.Space == "wp" || el.Space == "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing") {
		*result = append(*result, el)
	}
	for _, child := range el.ChildElements() {
		findInlines(child, result)
	}
}
