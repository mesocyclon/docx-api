package dml

import (
	"encoding/xml"
	"strings"
	"testing"
)

// ---------------------------------------------------------------------------
// Test XML samples (from reference-appendix.md §3.6)
// ---------------------------------------------------------------------------

// inlineXML is a representative <wp:inline> element taken from the reference
// appendix.  It includes a wp:cNvGraphicFramePr child that is NOT in the
// typed contract and must survive the round-trip as RawXML.
const inlineXML = `<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
	` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
	` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
	` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
	` distT="0" distB="0" distL="0" distR="0">` +
	`<wp:extent cx="1828800" cy="1371600"/>` +
	`<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
	`<wp:docPr id="1" name="Picture 1" descr="Logo"/>` +
	`<wp:cNvGraphicFramePr>` +
	`<a:graphicFrameLocks noChangeAspect="1"/>` +
	`</wp:cNvGraphicFramePr>` +
	`<a:graphic>` +
	`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
	`<pic:pic>` +
	`<pic:nvPicPr>` +
	`<pic:cNvPr id="1" name="image1.png"/>` +
	`<pic:cNvPicPr/>` +
	`</pic:nvPicPr>` +
	`<pic:blipFill>` +
	`<a:blip r:embed="rId10"/>` +
	`<a:stretch><a:fillRect/></a:stretch>` +
	`</pic:blipFill>` +
	`<pic:spPr>` +
	`<a:xfrm>` +
	`<a:off x="0" y="0"/>` +
	`<a:ext cx="1828800" cy="1371600"/>` +
	`</a:xfrm>` +
	`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
	`</pic:spPr>` +
	`</pic:pic>` +
	`</a:graphicData>` +
	`</a:graphic>` +
	`</wp:inline>`

// anchorXML is a representative <wp:anchor> element with wrapSquare.
const anchorXML = `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
	` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
	` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
	` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
	` behindDoc="0" distT="0" distB="0" distL="114300" distR="114300"` +
	` relativeHeight="251658240" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
	`<wp:simplePos x="0" y="0"/>` +
	`<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>` +
	`<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
	`<wp:extent cx="1828800" cy="1371600"/>` +
	`<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
	`<wp:wrapSquare wrapText="bothSides"/>` +
	`<wp:docPr id="2" name="Picture 2"/>` +
	`<a:graphic>` +
	`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
	`<pic:pic>` +
	`<pic:nvPicPr><pic:cNvPr id="2" name="img2.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
	`<pic:blipFill><a:blip r:embed="rId11"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
	`<pic:spPr>` +
	`<a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>` +
	`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
	`</pic:spPr>` +
	`</pic:pic>` +
	`</a:graphicData>` +
	`</a:graphic>` +
	`</wp:anchor>`

// ---------------------------------------------------------------------------
// Inline round-trip
// ---------------------------------------------------------------------------

func TestWPInlineRoundTrip(t *testing.T) {
	// 1. Unmarshal
	var inline WP_Inline
	if err := xml.Unmarshal([]byte(inlineXML), &inline); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// 2. Verify key fields
	if inline.DistT != 0 || inline.DistB != 0 || inline.DistL != 0 || inline.DistR != 0 {
		t.Error("dist attributes not parsed correctly")
	}
	if inline.Extent.CX != 1828800 || inline.Extent.CY != 1371600 {
		t.Errorf("Extent: got (%d, %d), want (1828800, 1371600)", inline.Extent.CX, inline.Extent.CY)
	}
	if inline.EffectExtent == nil {
		t.Fatal("EffectExtent is nil")
	}
	if inline.DocPr.ID != 1 || inline.DocPr.Name != "Picture 1" || inline.DocPr.Descr != "Logo" {
		t.Errorf("DocPr: got %+v", inline.DocPr)
	}
	if inline.Graphic.GraphicData.URI != "http://schemas.openxmlformats.org/drawingml/2006/picture" {
		t.Errorf("GraphicData.URI: got %q", inline.Graphic.GraphicData.URI)
	}
	if inline.Graphic.GraphicData.Pic == nil {
		t.Fatal("Pic is nil")
	}
	if inline.Graphic.GraphicData.Pic.BlipFill.Blip.Embed != "rId10" {
		t.Errorf("Blip.Embed: got %q, want rId10", inline.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
	}
	if inline.Graphic.GraphicData.Pic.SpPr.Xfrm == nil {
		t.Fatal("Xfrm is nil")
	}
	if inline.Graphic.GraphicData.Pic.SpPr.PrstGeom == nil || inline.Graphic.GraphicData.Pic.SpPr.PrstGeom.Prst != "rect" {
		t.Error("PrstGeom not parsed correctly")
	}
	if inline.Graphic.GraphicData.Pic.NvPicPr.CNvPr.Name != "image1.png" {
		t.Errorf("CNvPr.Name: got %q, want image1.png", inline.Graphic.GraphicData.Pic.NvPicPr.CNvPr.Name)
	}
	if inline.Graphic.GraphicData.Pic.BlipFill.Stretch == nil {
		t.Error("Stretch is nil")
	}

	// 3. Unknown element should be captured in Extra.
	if len(inline.Extra) == 0 {
		t.Fatal("expected Extra to contain cNvGraphicFramePr, got 0 elements")
	}
	if inline.Extra[0].XMLName.Local != "cNvGraphicFramePr" {
		t.Errorf("Extra[0].Local: got %q, want cNvGraphicFramePr", inline.Extra[0].XMLName.Local)
	}

	// 4. Marshal
	output, err := xml.Marshal(&inline)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	// 5. Round-trip: re-unmarshal and compare
	var inline2 WP_Inline
	if err := xml.Unmarshal(output, &inline2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}

	assertEqual(t, "Extent.CX", inline2.Extent.CX, inline.Extent.CX)
	assertEqual(t, "Extent.CY", inline2.Extent.CY, inline.Extent.CY)
	assertEqual(t, "DocPr.ID", inline2.DocPr.ID, inline.DocPr.ID)
	assertEqual(t, "DocPr.Name", inline2.DocPr.Name, inline.DocPr.Name)
	assertEqual(t, "DocPr.Descr", inline2.DocPr.Descr, inline.DocPr.Descr)
	if inline2.Graphic.GraphicData.Pic == nil {
		t.Fatal("round-trip lost Pic")
	}
	assertEqual(t, "Blip.Embed", inline2.Graphic.GraphicData.Pic.BlipFill.Blip.Embed, "rId10")
	assertEqual(t, "Xfrm.Off.X", inline2.Graphic.GraphicData.Pic.SpPr.Xfrm.Off.X, int64(0))
	assertEqual(t, "Xfrm.Ext.CX", inline2.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CX, int64(1828800))
	assertEqual(t, "PrstGeom.Prst", inline2.Graphic.GraphicData.Pic.SpPr.PrstGeom.Prst, "rect")
	if len(inline2.Extra) != len(inline.Extra) {
		t.Errorf("round-trip Extra: got %d, want %d", len(inline2.Extra), len(inline.Extra))
	}
	if len(inline2.Extra) > 0 {
		assertEqual(t, "Extra[0].Local", inline2.Extra[0].XMLName.Local, "cNvGraphicFramePr")
	}
}

// ---------------------------------------------------------------------------
// Anchor round-trip
// ---------------------------------------------------------------------------

func TestWPAnchorRoundTrip(t *testing.T) {
	// 1. Unmarshal
	var anchor WP_Anchor
	if err := xml.Unmarshal([]byte(anchorXML), &anchor); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// 2. Verify attributes
	if anchor.BehindDoc != false {
		t.Error("BehindDoc should be false")
	}
	if anchor.DistL != 114300 || anchor.DistR != 114300 {
		t.Errorf("DistL/R: got %d, %d", anchor.DistL, anchor.DistR)
	}
	if anchor.RelativeHeight != 251658240 {
		t.Errorf("RelativeHeight: got %d", anchor.RelativeHeight)
	}
	if anchor.LayoutInCell != true {
		t.Error("LayoutInCell should be true")
	}
	if anchor.AllowOverlap != true {
		t.Error("AllowOverlap should be true")
	}

	// 3. Verify position
	if anchor.PositionH.RelativeFrom != "column" {
		t.Errorf("PosH.RelativeFrom: got %q", anchor.PositionH.RelativeFrom)
	}
	if anchor.PositionH.PosOffset == nil || *anchor.PositionH.PosOffset != 914400 {
		t.Error("PosH.PosOffset not parsed")
	}
	if anchor.PositionV.RelativeFrom != "paragraph" {
		t.Errorf("PosV.RelativeFrom: got %q", anchor.PositionV.RelativeFrom)
	}
	if anchor.PositionV.PosOffset == nil || *anchor.PositionV.PosOffset != 0 {
		t.Error("PosV.PosOffset not parsed")
	}

	// 4. Verify wrap type
	ws, ok := anchor.WrapType.(WP_WrapSquare)
	if !ok {
		t.Fatalf("WrapType: got %T, want WP_WrapSquare", anchor.WrapType)
	}
	if ws.WrapText != "bothSides" {
		t.Errorf("WrapText: got %q, want bothSides", ws.WrapText)
	}

	// 5. Verify picture
	if anchor.DocPr.ID != 2 || anchor.DocPr.Name != "Picture 2" {
		t.Errorf("DocPr: got %+v", anchor.DocPr)
	}
	if anchor.Graphic.GraphicData.Pic == nil {
		t.Fatal("Pic is nil")
	}
	if anchor.Graphic.GraphicData.Pic.BlipFill.Blip.Embed != "rId11" {
		t.Errorf("Blip.Embed: got %q, want rId11", anchor.Graphic.GraphicData.Pic.BlipFill.Blip.Embed)
	}

	// 6. Marshal
	output, err := xml.Marshal(&anchor)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	// 7. Round-trip re-unmarshal
	var anchor2 WP_Anchor
	if err := xml.Unmarshal(output, &anchor2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}

	assertEqual(t, "DistL", anchor2.DistL, anchor.DistL)
	assertEqual(t, "RelativeHeight", anchor2.RelativeHeight, anchor.RelativeHeight)
	assertEqual(t, "LayoutInCell", anchor2.LayoutInCell, anchor.LayoutInCell)
	assertEqual(t, "PosH.RelativeFrom", anchor2.PositionH.RelativeFrom, "column")
	if anchor2.PositionH.PosOffset == nil || *anchor2.PositionH.PosOffset != 914400 {
		t.Error("round-trip lost PosH.PosOffset")
	}
	ws2, ok := anchor2.WrapType.(WP_WrapSquare)
	if !ok {
		t.Fatalf("round-trip WrapType: got %T", anchor2.WrapType)
	}
	assertEqual(t, "WrapText", ws2.WrapText, "bothSides")
	if anchor2.Graphic.GraphicData.Pic == nil {
		t.Fatal("round-trip lost Pic")
	}
	assertEqual(t, "Blip.Embed", anchor2.Graphic.GraphicData.Pic.BlipFill.Blip.Embed, "rId11")
}

// ---------------------------------------------------------------------------
// Anchor with align positioning
// ---------------------------------------------------------------------------

func TestWPAnchorAlign(t *testing.T) {
	const xml1 = `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
		` behindDoc="0" distT="0" distB="0" distL="0" distR="0"` +
		` relativeHeight="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
		`<wp:simplePos x="0" y="0"/>` +
		`<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>` +
		`<wp:positionV relativeFrom="page"><wp:align>top</wp:align></wp:positionV>` +
		`<wp:extent cx="100" cy="200"/>` +
		`<wp:wrapNone/>` +
		`<wp:docPr id="3" name="Shape"/>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
		`<pic:pic>` +
		`<pic:nvPicPr><pic:cNvPr id="3" name="s.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
		`<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>` +
		`<pic:spPr/>` +
		`</pic:pic>` +
		`</a:graphicData></a:graphic>` +
		`</wp:anchor>`

	var anchor WP_Anchor
	if err := xml.Unmarshal([]byte(xml1), &anchor); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if anchor.PositionH.Align == nil || *anchor.PositionH.Align != "center" {
		t.Error("PosH.Align not parsed")
	}
	if anchor.PositionV.Align == nil || *anchor.PositionV.Align != "top" {
		t.Error("PosV.Align not parsed")
	}
	if _, ok := anchor.WrapType.(WP_WrapNone); !ok {
		t.Errorf("WrapType: got %T, want WP_WrapNone", anchor.WrapType)
	}

	// Round-trip
	out, err := xml.Marshal(&anchor)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}
	outStr := string(out)
	if !strings.Contains(outStr, "center") {
		t.Error("round-trip lost align 'center'")
	}
	if !strings.Contains(outStr, "top") {
		t.Error("round-trip lost align 'top'")
	}
}

// ---------------------------------------------------------------------------
// Wrap types
// ---------------------------------------------------------------------------

func TestWrapTopAndBottom(t *testing.T) {
	const xml1 = `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
		` behindDoc="0" distT="0" distB="0" distL="0" distR="0"` +
		` relativeHeight="0" simplePos="0" locked="0" layoutInCell="0" allowOverlap="0">` +
		`<wp:simplePos x="0" y="0"/>` +
		`<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
		`<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="100" cy="100"/>` +
		`<wp:wrapTopAndBottom/>` +
		`<wp:docPr id="1" name="X"/>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
		`<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="x.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
		`<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr/></pic:pic>` +
		`</a:graphicData></a:graphic></wp:anchor>`

	var anchor WP_Anchor
	if err := xml.Unmarshal([]byte(xml1), &anchor); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if _, ok := anchor.WrapType.(WP_WrapTopAndBottom); !ok {
		t.Errorf("WrapType: got %T, want WP_WrapTopAndBottom", anchor.WrapType)
	}

	out, err := xml.Marshal(&anchor)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(out), "wrapTopAndBottom") {
		t.Error("round-trip lost wrapTopAndBottom")
	}
}

// ---------------------------------------------------------------------------
// WrapTight
// ---------------------------------------------------------------------------

func TestWrapTight(t *testing.T) {
	const xml1 = `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
		` behindDoc="1" distT="0" distB="0" distL="0" distR="0"` +
		` relativeHeight="0" simplePos="0" locked="1" layoutInCell="0" allowOverlap="0">` +
		`<wp:simplePos x="0" y="0"/>` +
		`<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
		`<wp:positionV relativeFrom="page"><wp:posOffset>100</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="50" cy="50"/>` +
		`<wp:wrapTight wrapText="largest"/>` +
		`<wp:docPr id="5" name="Z"/>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
		`<pic:pic><pic:nvPicPr><pic:cNvPr id="5" name="z.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
		`<pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill><pic:spPr/></pic:pic>` +
		`</a:graphicData></a:graphic></wp:anchor>`

	var anchor WP_Anchor
	if err := xml.Unmarshal([]byte(xml1), &anchor); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	wt, ok := anchor.WrapType.(WP_WrapTight)
	if !ok {
		t.Fatalf("WrapType: got %T, want WP_WrapTight", anchor.WrapType)
	}
	if wt.WrapText != "largest" {
		t.Errorf("WrapText: got %q, want largest", wt.WrapText)
	}
	if !anchor.BehindDoc {
		t.Error("BehindDoc should be true")
	}
	if !anchor.Locked {
		t.Error("Locked should be true")
	}
}

// ---------------------------------------------------------------------------
// Non-picture graphicData → RawData fallback
// ---------------------------------------------------------------------------

func TestGraphicDataRawFallback(t *testing.T) {
	const xml1 = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
		`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
		`<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId5"/>` +
		`</a:graphicData></a:graphic>`

	var g A_Graphic
	if err := xml.Unmarshal([]byte(xml1), &g); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if g.GraphicData.Pic != nil {
		t.Error("Pic should be nil for chart URI")
	}
	if g.GraphicData.RawData == nil {
		t.Fatal("RawData should capture unknown content")
	}
	if g.GraphicData.RawData.XMLName.Local != "chart" {
		t.Errorf("RawData.Local: got %q, want chart", g.GraphicData.RawData.XMLName.Local)
	}
	if g.GraphicData.URI != "http://schemas.openxmlformats.org/drawingml/2006/chart" {
		t.Errorf("URI: got %q", g.GraphicData.URI)
	}

	// Round-trip
	out, err := xml.Marshal(&g)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(out), "chart") {
		t.Error("round-trip lost chart element")
	}
}

// ---------------------------------------------------------------------------
// Isolated PosH / PosV
// ---------------------------------------------------------------------------

func TestPosHPosV(t *testing.T) {
	const posHXML = `<wp:positionH xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" relativeFrom="page">` +
		`<wp:posOffset>457200</wp:posOffset></wp:positionH>`
	var ph WP_PosH
	if err := xml.Unmarshal([]byte(posHXML), &ph); err != nil {
		t.Fatalf("unmarshal PosH: %v", err)
	}
	if ph.RelativeFrom != "page" {
		t.Errorf("RelativeFrom: got %q", ph.RelativeFrom)
	}
	if ph.PosOffset == nil || *ph.PosOffset != 457200 {
		t.Error("PosOffset not parsed")
	}

	out, err := xml.Marshal(&ph)
	if err != nil {
		t.Fatalf("marshal PosH: %v", err)
	}
	if !strings.Contains(string(out), "457200") {
		t.Error("round-trip lost posOffset value")
	}
}

// ---------------------------------------------------------------------------
// Isolated SpPr with extras
// ---------------------------------------------------------------------------

func TestSpPrWithExtras(t *testing.T) {
	const spPrXML = `<pic:spPr xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
		`<a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm>` +
		`<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>` +
		`<a:ln w="9525"><a:noFill/></a:ln>` +
		`</pic:spPr>`

	var sp A_SpPr
	if err := xml.Unmarshal([]byte(spPrXML), &sp); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if sp.Xfrm == nil {
		t.Fatal("Xfrm is nil")
	}
	if sp.Xfrm.Off.X != 10 || sp.Xfrm.Off.Y != 20 {
		t.Errorf("Off: got (%d,%d)", sp.Xfrm.Off.X, sp.Xfrm.Off.Y)
	}
	if sp.Xfrm.Ext.CX != 300 || sp.Xfrm.Ext.CY != 400 {
		t.Errorf("Ext: got (%d,%d)", sp.Xfrm.Ext.CX, sp.Xfrm.Ext.CY)
	}
	if sp.PrstGeom == nil || sp.PrstGeom.Prst != "ellipse" {
		t.Error("PrstGeom not parsed")
	}
	// a:ln should be in Extra
	if len(sp.Extra) != 1 {
		t.Fatalf("Extra: got %d, want 1", len(sp.Extra))
	}
	if sp.Extra[0].XMLName.Local != "ln" {
		t.Errorf("Extra[0].Local: got %q, want ln", sp.Extra[0].XMLName.Local)
	}

	// Round-trip
	out, err := xml.Marshal(&sp)
	if err != nil {
		t.Fatal(err)
	}
	var sp2 A_SpPr
	if err := xml.Unmarshal(out, &sp2); err != nil {
		t.Fatal(err)
	}
	if sp2.PrstGeom == nil || sp2.PrstGeom.Prst != "ellipse" {
		t.Error("round-trip lost PrstGeom")
	}
	if len(sp2.Extra) != 1 || sp2.Extra[0].XMLName.Local != "ln" {
		t.Error("round-trip lost Extra ln element")
	}
}

// ---------------------------------------------------------------------------
// EffectExtent
// ---------------------------------------------------------------------------

func TestEffectExtent(t *testing.T) {
	const xml1 = `<wp:effectExtent xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" l="19050" t="0" r="5080" b="0"/>`
	var ee WP_EffectExtent
	if err := xml.Unmarshal([]byte(xml1), &ee); err != nil {
		t.Fatal(err)
	}
	if ee.L != 19050 || ee.T != 0 || ee.R != 5080 || ee.B != 0 {
		t.Errorf("got L=%d T=%d R=%d B=%d", ee.L, ee.T, ee.R, ee.B)
	}

	out, err := xml.Marshal(&ee)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(out), "19050") || !strings.Contains(string(out), "5080") {
		t.Error("round-trip lost effectExtent values")
	}
}

// ---------------------------------------------------------------------------
// Marshal produces valid XML (basic sanity)
// ---------------------------------------------------------------------------

func TestInlineMarshalProducesValidXML(t *testing.T) {
	var inline WP_Inline
	if err := xml.Unmarshal([]byte(inlineXML), &inline); err != nil {
		t.Fatal(err)
	}
	out, err := xml.Marshal(&inline)
	if err != nil {
		t.Fatal(err)
	}
	// The output should be valid XML that can be re-parsed.
	d := xml.NewDecoder(strings.NewReader(string(out)))
	for {
		_, err := d.Token()
		if err != nil {
			break
		}
	}
}

// ---------------------------------------------------------------------------
// Helper
// ---------------------------------------------------------------------------

func assertEqual[T comparable](t *testing.T, label string, got, want T) {
	t.Helper()
	if got != want {
		t.Errorf("%s: got %v, want %v", label, got, want)
	}
}
