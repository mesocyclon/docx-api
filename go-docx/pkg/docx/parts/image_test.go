package parts

import (
	"crypto/sha1"
	"fmt"
	"math"
	"testing"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

func TestImagePartSHA1_Stable(t *testing.T) {
	blob := []byte("test image data")
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, blob, nil)

	h1 := ip.SHA1()
	h2 := ip.SHA1()
	if h1 != h2 {
		t.Errorf("SHA1 not stable: %q != %q", h1, h2)
	}

	// Verify against direct computation
	expected := fmt.Sprintf("%x", sha1.Sum(blob))
	if h1 != expected {
		t.Errorf("SHA1 = %q, want %q", h1, expected)
	}
}

func TestImagePartSHA1_SameBlob(t *testing.T) {
	blob := []byte("identical data")
	ip1 := NewImagePart("/word/media/image1.png", opc.CTPng, blob, nil)
	ip2 := NewImagePart("/word/media/image2.png", opc.CTPng, blob, nil)

	if ip1.SHA1() != ip2.SHA1() {
		t.Error("Same blob should produce same SHA1")
	}
}

func TestImagePartFilename(t *testing.T) {
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, nil, nil)
	// No explicit filename set — should fall back to partname ext
	got := ip.Filename()
	if got != "image.png" {
		t.Errorf("Filename = %q, want %q", got, "image.png")
	}

	ip.SetFilename("photo.jpg")
	got = ip.Filename()
	if got != "photo.jpg" {
		t.Errorf("Filename after set = %q, want %q", got, "photo.jpg")
	}
}

func TestImagePartDefaultCx(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	cx, err := ip.DefaultCx()
	if err != nil {
		t.Fatal(err)
	}
	expected := int64(math.Round(float64(100) / float64(96) * 914400))
	if cx != expected {
		t.Errorf("DefaultCx = %d, want %d", cx, expected)
	}
}

func TestImagePartDefaultCy_UsesHorzDpi(t *testing.T) {
	// Python uses horz_dpi for cy, not vert_dpi!
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 300, "",
	)
	cy, err := ip.DefaultCy()
	if err != nil {
		t.Fatal(err)
	}
	// Uses horz_dpi (96), not vert_dpi (300)
	expected := int64(math.Round(914400 * float64(200) / float64(96)))
	if cy != expected {
		t.Errorf("DefaultCy = %d, want %d (should use horz_dpi)", cy, expected)
	}
}

func TestImagePartDefaultCx_NoDPI(t *testing.T) {
	ip := NewImagePart("/word/media/image1.png", opc.CTPng, nil, nil)
	_, err := ip.DefaultCx()
	if err == nil {
		t.Error("expected error for image with no DPI")
	}
}

func TestScaledDimensions_BothNil(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	cx, cy, err := ip.ScaledDimensions(nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	nativeW, _ := ip.DefaultCx()
	nativeH, _ := ip.DefaultCy()
	if cx != nativeW || cy != nativeH {
		t.Errorf("ScaledDimensions(nil,nil) = (%d,%d), want (%d,%d)", cx, cy, nativeW, nativeH)
	}
}

func TestScaledDimensions_WidthOnly(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	w := int64(457200)
	cx, cy, err := ip.ScaledDimensions(&w, nil)
	if err != nil {
		t.Fatal(err)
	}
	if cx != w {
		t.Errorf("cx = %d, want %d", cx, w)
	}
	nativeW, _ := ip.DefaultCx()
	nativeH, _ := ip.DefaultCy()
	expectedH := int64(math.Round(float64(nativeH) * float64(w) / float64(nativeW)))
	if cy != expectedH {
		t.Errorf("cy = %d, want %d", cy, expectedH)
	}
}

func TestScaledDimensions_HeightOnly(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	h := int64(914400)
	cx, cy, err := ip.ScaledDimensions(nil, &h)
	if err != nil {
		t.Fatal(err)
	}
	if cy != h {
		t.Errorf("cy = %d, want %d", cy, h)
	}
	nativeW, _ := ip.DefaultCx()
	nativeH, _ := ip.DefaultCy()
	expectedW := int64(math.Round(float64(nativeW) * float64(h) / float64(nativeH)))
	if cx != expectedW {
		t.Errorf("cx = %d, want %d", cx, expectedW)
	}
}

func TestScaledDimensions_BothSpecified(t *testing.T) {
	ip := NewImagePartWithMeta(
		"/word/media/image1.png", opc.CTPng, nil,
		100, 200, 96, 96, "",
	)
	w, h := int64(111), int64(222)
	cx, cy, err := ip.ScaledDimensions(&w, &h)
	if err != nil {
		t.Fatal(err)
	}
	if cx != w || cy != h {
		t.Errorf("ScaledDimensions(&111,&222) = (%d,%d), want (111,222)", cx, cy)
	}
}
