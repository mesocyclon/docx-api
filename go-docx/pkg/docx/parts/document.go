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
	// cached related parts — mirrors Python lazyproperty / try-except KeyError pattern
	stylesPart    *StylesPart
	numberingPart *NumberingPart
	settingsPart  *SettingsPart
	commentsPart  *CommentsPart
}

// NewDocumentPart creates a DocumentPart wrapping the given XmlPart.
func NewDocumentPart(xp *opc.XmlPart) *DocumentPart {
	dp := &DocumentPart{
		StoryPart: StoryPart{XmlPart: xp},
	}
	// The document part is its own document part.
	dp.StoryPart.docPart = dp
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
// Mirrors Python DocumentPart.add_header_part.
func (dp *DocumentPart) AddHeaderPart() (*HeaderPart, string, error) {
	pkg := dp.Package()
	if pkg == nil {
		return nil, "", fmt.Errorf("parts: document part has no package")
	}
	hp, err := NewHeaderPart(pkg)
	if err != nil {
		return nil, "", fmt.Errorf("parts: creating header part: %w", err)
	}
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
func (dp *DocumentPart) DropHeaderPart(rId string) {
	dp.DropRel(rId)
}

// HeaderPartByRID returns the HeaderPart related by the given rId.
//
// Mirrors Python DocumentPart.header_part(rId).
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
// Mirrors Python DocumentPart.footer_part(rId).
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
// StylesPart
// --------------------------------------------------------------------------

// StylesPart returns the StylesPart for this document, creating a default
// one if not present. The result is cached.
//
// Mirrors Python DocumentPart._styles_part property.
func (dp *DocumentPart) StylesPart() (*StylesPart, error) {
	if dp.stylesPart != nil {
		return dp.stylesPart, nil
	}
	rel, err := dp.Rels().GetByRelType(opc.RTStyles)
	if err == nil && rel.TargetPart != nil {
		sp, ok := rel.TargetPart.(*StylesPart)
		if ok {
			dp.stylesPart = sp
			return sp, nil
		}
	}
	// Not found — create default
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
	dp.stylesPart = sp
	return sp, nil
}

// Styles returns the CT_Styles element from the styles part.
//
// Mirrors Python DocumentPart.styles property.
func (dp *DocumentPart) Styles() (*oxml.CT_Styles, error) {
	sp, err := dp.StylesPart()
	if err != nil {
		return nil, err
	}
	return sp.Styles()
}

// --------------------------------------------------------------------------
// NumberingPart
// --------------------------------------------------------------------------

// NumberingPart returns the NumberingPart for this document. Unlike styles
// and settings, numbering does not auto-create a default part in the Python
// source (NumberingPart.new() raises NotImplementedError). We only resolve
// existing relationships.
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
// SettingsPart
// --------------------------------------------------------------------------

// SettingsPart returns the SettingsPart for this document, creating a
// default one if not present. The result is cached.
//
// Mirrors Python DocumentPart._settings_part property.
func (dp *DocumentPart) SettingsPart() (*SettingsPart, error) {
	if dp.settingsPart != nil {
		return dp.settingsPart, nil
	}
	rel, err := dp.Rels().GetByRelType(opc.RTSettings)
	if err == nil && rel.TargetPart != nil {
		sp, ok := rel.TargetPart.(*SettingsPart)
		if ok {
			dp.settingsPart = sp
			return sp, nil
		}
	}
	// Not found — create default
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
	dp.settingsPart = sp
	return sp, nil
}

// Settings returns the CT_Settings element from the settings part.
//
// Mirrors Python DocumentPart.settings property.
func (dp *DocumentPart) Settings() (*oxml.CT_Settings, error) {
	sp, err := dp.SettingsPart()
	if err != nil {
		return nil, err
	}
	return sp.SettingsElement()
}

// --------------------------------------------------------------------------
// CommentsPart
// --------------------------------------------------------------------------

// CommentsPart returns the CommentsPart for this document, creating a
// default one if not present. The result is cached.
//
// Mirrors Python DocumentPart._comments_part property.
func (dp *DocumentPart) CommentsPart() (*CommentsPart, error) {
	if dp.commentsPart != nil {
		return dp.commentsPart, nil
	}
	rel, err := dp.Rels().GetByRelType(opc.RTComments)
	if err == nil && rel.TargetPart != nil {
		cp, ok := rel.TargetPart.(*CommentsPart)
		if ok {
			dp.commentsPart = cp
			return cp, nil
		}
	}
	// Not found — create default
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
	dp.commentsPart = cp
	return cp, nil
}

// CommentsElement returns the CT_Comments element from the comments part.
//
// Mirrors Python DocumentPart.comments property (partially — domain
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
// Mirrors Python DocumentPart.core_properties (delegates to package).
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

// GetStyle returns the CT_Style matching styleID and styleType.
// If styleID is nil the default style for styleType is returned.
//
// Mirrors Python DocumentPart.get_style.
func (dp *DocumentPart) GetStyle(styleID *string, styleType enum.WdStyleType) (*etree.Element, error) {
	ss, err := dp.Styles()
	if err != nil {
		return nil, err
	}
	if styleID == nil {
		defStyle := ss.DefaultFor(styleType)
		if defStyle == nil {
			return nil, nil
		}
		return defStyle.E, nil
	}
	s := ss.GetByID(*styleID)
	if s != nil {
		return s.E, nil
	}
	// Fall back to default for the style type.
	defStyle := ss.DefaultFor(styleType)
	if defStyle == nil {
		return nil, nil
	}
	return defStyle.E, nil
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
