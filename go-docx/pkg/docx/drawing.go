package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Drawing is a container for a DrawingML object within a run.
//
// Mirrors Python Drawing(Parented) from docx/drawing/__init__.py.
type Drawing struct {
	drawing *oxml.CT_Drawing
	part    *parts.StoryPart
}

// NewDrawing creates a new Drawing proxy.
func NewDrawing(drawing *oxml.CT_Drawing, part *parts.StoryPart) *Drawing {
	return &Drawing{drawing: drawing, part: part}
}

// HasPicture returns true when the drawing contains an embedded picture.
// A drawing can also contain a chart, SmartArt, or drawing canvas.
//
// Checks for both inline and floating pictures:
//
//	wp:inline/a:graphic/a:graphicData/pic:pic
//	wp:anchor/a:graphic/a:graphicData/pic:pic
//
// Mirrors Python Drawing.has_picture.
func (d *Drawing) HasPicture() bool {
	for _, child := range d.drawing.E.ChildElements() {
		if child.Tag == "inline" || child.Tag == "anchor" {
			if findPicInGraphicData(child) {
				return true
			}
		}
	}
	return false
}

// ImagePart returns the ImagePart for the embedded picture in this drawing.
// Returns an error if the drawing does not contain a picture.
//
// Mirrors Python Drawing.image which returns image_part.image.
func (d *Drawing) ImagePart() (*parts.ImagePart, error) {
	rId := d.pictureRId()
	if rId == "" {
		return nil, fmt.Errorf("docx: drawing does not contain a picture")
	}
	rels := d.part.Rels()
	if rels == nil {
		return nil, fmt.Errorf("docx: drawing part has no relationships")
	}
	relParts := rels.RelatedParts()
	p, ok := relParts[rId]
	if !ok {
		return nil, fmt.Errorf("docx: no related part for rId %q", rId)
	}
	ip, ok := p.(*parts.ImagePart)
	if !ok {
		return nil, fmt.Errorf("docx: related part for rId %q is not an ImagePart", rId)
	}
	return ip, nil
}

// pictureRId finds the r:embed attribute on a:blip inside pic:blipFill.
// Walks: drawing → inline/anchor → graphic → graphicData → pic → blipFill → blip → @r:embed.
//
// Mirrors Python: self._drawing.xpath(".//pic:blipFill/a:blip/@r:embed")
func (d *Drawing) pictureRId() string {
	for _, child := range d.drawing.E.ChildElements() {
		if child.Tag == "inline" || child.Tag == "anchor" {
			if rId := findBlipRId(child); rId != "" {
				return rId
			}
		}
	}
	return ""
}

// findBlipRId recursively searches for pic:blipFill/a:blip/@r:embed.
func findBlipRId(el *etree.Element) string {
	for _, child := range el.ChildElements() {
		if child.Tag == "blipFill" {
			for _, blipChild := range child.ChildElements() {
				if blipChild.Tag == "blip" {
					for _, attr := range blipChild.Attr {
						if attr.Key == "embed" && (attr.Space == "r" || attr.Space == "") {
							return attr.Value
						}
					}
				}
			}
		}
		// Recurse into children (graphic → graphicData → pic → ...)
		if rId := findBlipRId(child); rId != "" {
			return rId
		}
	}
	return ""
}

// findPicInGraphicData walks child → a:graphic → a:graphicData → pic:pic.
func findPicInGraphicData(el *etree.Element) bool {
	for _, graphic := range el.ChildElements() {
		if graphic.Tag == "graphic" {
			for _, graphicData := range graphic.ChildElements() {
				if graphicData.Tag == "graphicData" {
					for _, pic := range graphicData.ChildElements() {
						if pic.Tag == "pic" {
							return true
						}
					}
				}
			}
		}
	}
	return false
}

// CT_Drawing returns the underlying oxml element.
func (d *Drawing) CT_Drawing() *oxml.CT_Drawing { return d.drawing }
