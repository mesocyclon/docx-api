package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// drawing_test.go — Drawing (Batch 1)
// Mirrors Python: tests/test_drawing.py
// -----------------------------------------------------------------------

// Mirrors Python: it_knows_when_it_contains_a_Picture
func TestDrawing_HasPicture(t *testing.T) {
	tests := []struct {
		name     string
		xml      string
		expected bool
	}{
		{
			"has_picture",
			`<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
				xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
				xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
				<wp:inline>
					<a:graphic>
						<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
							<pic:pic/>
						</a:graphicData>
					</a:graphic>
				</wp:inline>
			</w:drawing>`,
			true,
		},
		{
			"no_picture_chart",
			`<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
				xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
				<wp:inline>
					<a:graphic>
						<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
					</a:graphic>
				</wp:inline>
			</w:drawing>`,
			false,
		},
		{
			"empty_drawing",
			`<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
			false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			el, err := oxml.ParseXml([]byte(tt.xml))
			if err != nil {
				t.Fatal(err)
			}
			d := &oxml.CT_Drawing{Element: oxml.WrapElement(el)}
			drawing := newDrawing(d, nil)
			if got := drawing.HasPicture(); got != tt.expected {
				t.Errorf("HasPicture() = %v, want %v", got, tt.expected)
			}
		})
	}
}

// Mirrors Python: it_provides_access_to_the_image
// Tests both error paths: no picture, and no part.
func TestDrawing_ImagePart_NoPicture(t *testing.T) {
	// Drawing with a chart (no pic:pic) → ImagePart should error "does not contain a picture"
	xml := `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
		<wp:inline>
			<a:graphic>
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
			</a:graphic>
		</wp:inline>
	</w:drawing>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	d := &oxml.CT_Drawing{Element: oxml.WrapElement(el)}
	drawing := newDrawing(d, nil)

	_, err = drawing.ImagePart()
	if err == nil {
		t.Error("expected error for ImagePart on drawing without picture")
	}
}

func TestDrawing_ImagePart_NoPart(t *testing.T) {
	// Drawing WITH a picture but nil part → should error gracefully, not panic
	xml := `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<wp:inline>
			<a:graphic>
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:pic>
						<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		</wp:inline>
	</w:drawing>`
	el, err := oxml.ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	d := &oxml.CT_Drawing{Element: oxml.WrapElement(el)}
	drawing := newDrawing(d, nil) // nil part → should error, not panic

	_, err = drawing.ImagePart()
	if err == nil {
		t.Error("expected error for ImagePart with nil part")
	}
}
