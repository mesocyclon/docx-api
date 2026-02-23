// Package parts implements WML-specific part types (document, styles, headers, etc.)
// that extend the generic OPC part infrastructure.
package parts

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/opc"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// StoryPart is the base for document-body, header, and footer parts.
// A story part is one that can contain textual content. These all share
// content behaviours like paragraphs, tables, and images.
//
// Mirrors Python StoryPart(XmlPart).
type StoryPart struct {
	*opc.XmlPart
	docPart *DocumentPart // cached, mirrors Python lazyproperty _document_part
}

// NewStoryPart creates a StoryPart wrapping the given XmlPart.
func NewStoryPart(xp *opc.XmlPart) *StoryPart {
	return &StoryPart{XmlPart: xp}
}

// GetOrAddImage returns (rId, image) for the image read from r.
// The image-part is reused if an identical image (by SHA1) already exists
// in the package.
//
// Mirrors Python StoryPart.get_or_add_image.
func (sp *StoryPart) GetOrAddImage(wmlPkg *WmlPackage) (string, *ImagePart, error) {
	// NOTE: the actual io.ReadSeeker-based version will be added in MR-10
	// when the image.Image type is available. This is a placeholder that
	// the DocumentPart and callers will route through WmlPackage.
	return "", nil, fmt.Errorf("parts: GetOrAddImage requires image layer (MR-10)")
}

// GetOrAddImagePart adds (or reuses) an image part from a WmlPackage and
// returns the relationship ID from this story part to the image part.
//
// Mirrors Python StoryPart.get_or_add_image (relationship wiring portion).
func (sp *StoryPart) GetOrAddImagePart(imgPart *ImagePart) string {
	rel := sp.Rels().GetOrAdd(opc.RTImage, imgPart)
	return rel.RID
}

// NewPicInline creates a new CT_Inline element containing the image specified
// by the given image part, scaled to the given width/height.
//
// Mirrors Python StoryPart.new_pic_inline.
func (sp *StoryPart) NewPicInline(imgPart *ImagePart, rId string, width, height *int64) (*oxml.CT_Inline, error) {
	cx, cy, err := imgPart.ScaledDimensions(width, height)
	if err != nil {
		return nil, fmt.Errorf("parts: computing scaled dimensions: %w", err)
	}
	shapeID := sp.NextID()
	filename := imgPart.Filename()
	return oxml.NewPicInline(shapeID, rId, filename, cx, cy)
}

// NextID returns the next available positive integer id value in this story
// XML document. The value is determined by incrementing the maximum existing
// id value. Gaps in the existing id sequence are not filled.
//
// Mirrors Python StoryPart.next_id.
func (sp *StoryPart) NextID() int {
	el := sp.Element()
	if el == nil {
		return 1
	}
	maxID := 0
	collectMaxID(el, &maxID)
	return maxID + 1
}

// collectMaxID walks the element tree collecting all @id attributes that are
// purely numeric digits, tracking the maximum value found.
func collectMaxID(el *etree.Element, maxID *int) {
	for _, attr := range el.Attr {
		if attr.Key == "id" && attr.Space == "" {
			if isDigits(attr.Value) {
				if v, err := strconv.Atoi(attr.Value); err == nil && v > *maxID {
					*maxID = v
				}
			}
		}
	}
	for _, child := range el.ChildElements() {
		collectMaxID(child, maxID)
	}
}

// isDigits returns true if s is non-empty and consists only of ASCII digits.
func isDigits(s string) bool {
	if s == "" {
		return false
	}
	for _, c := range s {
		if c < '0' || c > '9' {
			return false
		}
	}
	return true
}

// DocumentPart returns the main DocumentPart for the package this story part
// belongs to. The result is cached after the first call.
//
// Mirrors Python StoryPart._document_part (lazyproperty).
func (sp *StoryPart) DocumentPart() (*DocumentPart, error) {
	if sp.docPart != nil {
		return sp.docPart, nil
	}
	pkg := sp.Package()
	if pkg == nil {
		return nil, fmt.Errorf("parts: story part has no package")
	}
	mainPart, err := pkg.MainDocumentPart()
	if err != nil {
		return nil, fmt.Errorf("parts: resolving document part: %w", err)
	}
	dp, ok := mainPart.(*DocumentPart)
	if !ok {
		return nil, fmt.Errorf("parts: main document part is %T, want *DocumentPart", mainPart)
	}
	sp.docPart = dp
	return dp, nil
}

// DropRel removes the relationship identified by rId if its reference count
// in this part's XML is less than 2. This prevents removing relationships
// that are still referenced elsewhere in the XML.
//
// Mirrors Python Part.drop_rel + XmlPart._rel_ref_count.
func (sp *StoryPart) DropRel(rId string) {
	if sp.relRefCount(rId) < 2 {
		sp.Rels().Delete(rId)
	}
}

// relRefCount returns the count of references to rId in this part's XML.
// Mirrors Python XmlPart._rel_ref_count which counts //@r:id occurrences.
func (sp *StoryPart) relRefCount(rId string) int {
	el := sp.Element()
	if el == nil {
		return 0
	}
	count := 0
	countRIdRefs(el, rId, &count)
	return count
}

// countRIdRefs recursively counts attributes named r:id (or {relationship-ns}id)
// with the given value.
func countRIdRefs(el *etree.Element, rId string, count *int) {
	for _, attr := range el.Attr {
		if attr.Key == "id" && isRelNS(attr.Space) && attr.Value == rId {
			*count++
		}
	}
	for _, child := range el.ChildElements() {
		countRIdRefs(child, rId, count)
	}
}

// isRelNS returns true if the namespace prefix or URI matches the OFC
// relationships namespace used for r:id attributes.
func isRelNS(space string) bool {
	return space == "r" ||
		strings.Contains(space, "officeDocument/2006/relationships")
}
