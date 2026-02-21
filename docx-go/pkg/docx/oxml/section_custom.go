package oxml

import (
	"github.com/beevik/etree"
	"github.com/user/go-docx/pkg/docx/enum"
)

// ===========================================================================
// CT_SectPr — custom methods
// ===========================================================================

// Clone returns a deep copy of this sectPr element with all rsid attributes removed.
func (sp *CT_SectPr) Clone() *CT_SectPr {
	copied := sp.E.Copy()
	// Remove rsid* attributes
	var toRemove []string
	for _, attr := range copied.Attr {
		key := attr.Key
		if len(key) >= 4 && key[:4] == "rsid" {
			toRemove = append(toRemove, attr.FullKey())
		}
		if attr.Space != "" {
			qKey := attr.Space + ":" + attr.Key
			if len(attr.Key) >= 4 && attr.Key[:4] == "rsid" {
				toRemove = append(toRemove, qKey)
			}
		}
	}
	for _, k := range toRemove {
		copied.RemoveAttr(k)
	}
	return &CT_SectPr{Element{E: copied}}
}

// --- Page size ---

// PageWidth returns the page width in twips from pgSz/@w:w, or nil.
func (sp *CT_SectPr) PageWidth() *int {
	pgSz := sp.PgSz()
	if pgSz == nil {
		return nil
	}
	v := pgSz.W()
	if v == 0 {
		return nil
	}
	return &v
}

// SetPageWidth sets the page width in twips.
func (sp *CT_SectPr) SetPageWidth(twips *int) {
	if twips == nil {
		pgSz := sp.PgSz()
		if pgSz != nil {
			pgSz.SetW(0)
		}
		return
	}
	sp.GetOrAddPgSz().SetW(*twips)
}

// PageHeight returns the page height in twips from pgSz/@w:h, or nil.
func (sp *CT_SectPr) PageHeight() *int {
	pgSz := sp.PgSz()
	if pgSz == nil {
		return nil
	}
	v := pgSz.H()
	if v == 0 {
		return nil
	}
	return &v
}

// SetPageHeight sets the page height in twips.
func (sp *CT_SectPr) SetPageHeight(twips *int) {
	if twips == nil {
		pgSz := sp.PgSz()
		if pgSz != nil {
			pgSz.SetH(0)
		}
		return
	}
	sp.GetOrAddPgSz().SetH(*twips)
}

// --- Orientation ---

// Orientation returns the page orientation. Defaults to PORTRAIT when not present.
func (sp *CT_SectPr) Orientation() enum.WdOrientation {
	pgSz := sp.PgSz()
	if pgSz == nil {
		return enum.WdOrientationPortrait
	}
	v := pgSz.Orient()
	if v == enum.WdOrientation(0) {
		return enum.WdOrientationPortrait
	}
	return v
}

// SetOrientation sets the page orientation.
func (sp *CT_SectPr) SetOrientation(v enum.WdOrientation) {
	pgSz := sp.GetOrAddPgSz()
	if v == enum.WdOrientationPortrait {
		pgSz.SetOrient(enum.WdOrientation(0)) // removes attr, defaulting to portrait
	} else {
		pgSz.SetOrient(v)
	}
}

// --- Start type ---

// StartType returns the section start type. Defaults to NEW_PAGE when not present.
func (sp *CT_SectPr) StartType() enum.WdSectionStart {
	t := sp.Type()
	if t == nil {
		return enum.WdSectionStartNewPage
	}
	// Check if val attribute is actually present
	_, ok := t.GetAttr("w:val")
	if !ok {
		return enum.WdSectionStartNewPage
	}
	return t.Val()
}

// SetStartType sets the section start type. Passing WdSectionStartNewPage removes
// the type element (since NEW_PAGE is the default).
func (sp *CT_SectPr) SetStartType(v enum.WdSectionStart) {
	if v == enum.WdSectionStartNewPage {
		sp.RemoveType()
		return
	}
	t := sp.GetOrAddType()
	// Use SetAttr directly because the generated SetVal treats zero (Continuous)
	// as "remove attribute", but we actually need to write w:val="continuous".
	t.SetAttr("w:val", v.ToXml())
}

// --- Title page ---

// TitlePgVal returns true if the first page has different header/footer.
func (sp *CT_SectPr) TitlePgVal() bool {
	tp := sp.TitlePg()
	if tp == nil {
		return false
	}
	return tp.Val()
}

// SetTitlePgVal sets the titlePg flag. Passing false removes the element.
func (sp *CT_SectPr) SetTitlePgVal(v bool) {
	if !v {
		sp.RemoveTitlePg()
		return
	}
	sp.GetOrAddTitlePg().SetVal(true)
}

// --- Margins ---

// TopMargin returns the top margin in twips, or nil if not present.
func (sp *CT_SectPr) TopMargin() *int {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil
	}
	v := pgMar.Top()
	if v == 0 {
		return nil
	}
	return &v
}

// SetTopMargin sets the top margin in twips. Passing nil removes the attribute.
func (sp *CT_SectPr) SetTopMargin(twips *int) {
	pgMar := sp.GetOrAddPgMar()
	if twips == nil {
		pgMar.SetTop(0)
	} else {
		pgMar.SetTop(*twips)
	}
}

// BottomMargin returns the bottom margin in twips, or nil.
func (sp *CT_SectPr) BottomMargin() *int {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil
	}
	v := pgMar.Bottom()
	if v == 0 {
		return nil
	}
	return &v
}

// SetBottomMargin sets the bottom margin in twips.
func (sp *CT_SectPr) SetBottomMargin(twips *int) {
	pgMar := sp.GetOrAddPgMar()
	if twips == nil {
		pgMar.SetBottom(0)
	} else {
		pgMar.SetBottom(*twips)
	}
}

// LeftMargin returns the left margin in twips, or nil.
func (sp *CT_SectPr) LeftMargin() *int {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil
	}
	v := pgMar.Left()
	if v == 0 {
		return nil
	}
	return &v
}

// SetLeftMargin sets the left margin in twips.
func (sp *CT_SectPr) SetLeftMargin(twips *int) {
	pgMar := sp.GetOrAddPgMar()
	if twips == nil {
		pgMar.SetLeft(0)
	} else {
		pgMar.SetLeft(*twips)
	}
}

// RightMargin returns the right margin in twips, or nil.
func (sp *CT_SectPr) RightMargin() *int {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil
	}
	v := pgMar.Right()
	if v == 0 {
		return nil
	}
	return &v
}

// SetRightMargin sets the right margin in twips.
func (sp *CT_SectPr) SetRightMargin(twips *int) {
	pgMar := sp.GetOrAddPgMar()
	if twips == nil {
		pgMar.SetRight(0)
	} else {
		pgMar.SetRight(*twips)
	}
}

// HeaderMargin returns the header distance from top edge in twips, or nil.
func (sp *CT_SectPr) HeaderMargin() *int {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil
	}
	v := pgMar.Header()
	if v == 0 {
		return nil
	}
	return &v
}

// SetHeaderMargin sets the header margin in twips.
func (sp *CT_SectPr) SetHeaderMargin(twips *int) {
	pgMar := sp.GetOrAddPgMar()
	if twips == nil {
		pgMar.SetHeader(0)
	} else {
		pgMar.SetHeader(*twips)
	}
}

// FooterMargin returns the footer distance from bottom edge in twips, or nil.
func (sp *CT_SectPr) FooterMargin() *int {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil
	}
	v := pgMar.Footer()
	if v == 0 {
		return nil
	}
	return &v
}

// SetFooterMargin sets the footer margin in twips.
func (sp *CT_SectPr) SetFooterMargin(twips *int) {
	pgMar := sp.GetOrAddPgMar()
	if twips == nil {
		pgMar.SetFooter(0)
	} else {
		pgMar.SetFooter(*twips)
	}
}

// GutterMargin returns the gutter in twips, or nil.
func (sp *CT_SectPr) GutterMargin() *int {
	pgMar := sp.PgMar()
	if pgMar == nil {
		return nil
	}
	v := pgMar.Gutter()
	if v == 0 {
		return nil
	}
	return &v
}

// SetGutterMargin sets the gutter in twips.
func (sp *CT_SectPr) SetGutterMargin(twips *int) {
	pgMar := sp.GetOrAddPgMar()
	if twips == nil {
		pgMar.SetGutter(0)
	} else {
		pgMar.SetGutter(*twips)
	}
}

// --- Header/Footer references ---

// AddHeaderRef adds a headerReference with the given type and relationship ID.
func (sp *CT_SectPr) AddHeaderRef(hfType enum.WdHeaderFooterIndex, rId string) *CT_HdrFtrRef {
	ref := sp.AddHeaderReference()
	ref.SetType(hfType)
	ref.SetRId(rId)
	return ref
}

// AddFooterRef adds a footerReference with the given type and relationship ID.
func (sp *CT_SectPr) AddFooterRef(hfType enum.WdHeaderFooterIndex, rId string) *CT_HdrFtrRef {
	ref := sp.AddFooterReference()
	ref.SetType(hfType)
	ref.SetRId(rId)
	return ref
}

// GetHeaderRef returns the headerReference of the given type, or nil.
func (sp *CT_SectPr) GetHeaderRef(hfType enum.WdHeaderFooterIndex) *CT_HdrFtrRef {
	xmlVal := hfType.ToXml()
	for _, ref := range sp.HeaderReferenceList() {
		v, ok := ref.GetAttr("w:type")
		if ok && v == xmlVal {
			return ref
		}
	}
	return nil
}

// GetFooterRef returns the footerReference of the given type, or nil.
func (sp *CT_SectPr) GetFooterRef(hfType enum.WdHeaderFooterIndex) *CT_HdrFtrRef {
	xmlVal := hfType.ToXml()
	for _, ref := range sp.FooterReferenceList() {
		v, ok := ref.GetAttr("w:type")
		if ok && v == xmlVal {
			return ref
		}
	}
	return nil
}

// RemoveHeaderRef removes the headerReference of the given type and returns its rId.
// Returns "" if not found.
func (sp *CT_SectPr) RemoveHeaderRef(hfType enum.WdHeaderFooterIndex) string {
	ref := sp.GetHeaderRef(hfType)
	if ref == nil {
		return ""
	}
	rId, _ := ref.RId()
	sp.E.RemoveChild(ref.E)
	return rId
}

// RemoveFooterRef removes the footerReference of the given type and returns its rId.
// Returns "" if not found.
func (sp *CT_SectPr) RemoveFooterRef(hfType enum.WdHeaderFooterIndex) string {
	ref := sp.GetFooterRef(hfType)
	if ref == nil {
		return ""
	}
	rId, _ := ref.RId()
	sp.E.RemoveChild(ref.E)
	return rId
}

// PrecedingSectPr returns the sectPr immediately preceding this one, or nil.
// Searches via preceding-sibling within the w:body, accounting for both
// paragraph-based sectPr (w:p/w:pPr/w:sectPr) and body-based sectPr (w:body/w:sectPr).
func (sp *CT_SectPr) PrecedingSectPr() *CT_SectPr {
	// Determine if this is body-based or pPr-based
	parent := sp.E.Parent()
	if parent == nil {
		return nil
	}

	// Collect all sectPr in the body to find this one's predecessor
	var body *etree.Element
	if parent.Space == "w" && parent.Tag == "body" {
		body = parent
	} else if parent.Space == "w" && parent.Tag == "pPr" {
		p := parent.Parent()
		if p != nil {
			body = p.Parent()
		}
	}
	if body == nil {
		return nil
	}

	// Gather all sectPr elements in document order
	var allSectPrs []*CT_SectPr
	for _, child := range body.ChildElements() {
		// Check p/pPr/sectPr
		if child.Space == "w" && child.Tag == "p" {
			for _, pChild := range child.ChildElements() {
				if pChild.Space == "w" && pChild.Tag == "pPr" {
					for _, ppChild := range pChild.ChildElements() {
						if ppChild.Space == "w" && ppChild.Tag == "sectPr" {
							allSectPrs = append(allSectPrs, &CT_SectPr{Element{E: ppChild}})
						}
					}
				}
			}
		}
		// Check body/sectPr
		if child.Space == "w" && child.Tag == "sectPr" {
			allSectPrs = append(allSectPrs, &CT_SectPr{Element{E: child}})
		}
	}

	for i, s := range allSectPrs {
		if s.E == sp.E && i > 0 {
			return allSectPrs[i-1]
		}
	}
	return nil
}

// ===========================================================================
// CT_HdrFtr — custom methods
// ===========================================================================

// InnerContentElements returns all w:p and w:tbl direct children in document order.
func (hf *CT_HdrFtr) InnerContentElements() []interface{} {
	var result []interface{}
	for _, child := range hf.E.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			result = append(result, &CT_P{Element{E: child}})
		} else if child.Space == "w" && child.Tag == "tbl" {
			result = append(result, &CT_Tbl{Element{E: child}})
		}
	}
	return result
}

