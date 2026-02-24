package parts

import (
	"crypto/sha1"
	"fmt"
	"math"

	"github.com/vortex/go-docx/pkg/docx/image"
	"github.com/vortex/go-docx/pkg/docx/opc"
)

// ImagePart stores an image as a binary blob. Unlike most parts, it does not
// contain XML — it embeds opc.BasePart directly.
//
// Mirrors Python ImagePart(Part).
type ImagePart struct {
	*opc.BasePart
	sha1Hash   string // lazy, "" until first SHA1() call
	metaLoaded bool   // true once dimensions/DPI have been parsed

	// Image metadata — populated lazily from blob or explicitly via SetImageMeta.
	pxWidth  int
	pxHeight int
	horzDpi  int
	vertDpi  int
	filename string // original filename if known
}

// NewImagePart creates an ImagePart with the given blob data.
func NewImagePart(partName opc.PackURI, contentType string, blob []byte, pkg *opc.OpcPackage) *ImagePart {
	return &ImagePart{
		BasePart: opc.NewBasePart(partName, contentType, blob, pkg),
	}
}

// NewImagePartWithMeta creates an ImagePart with full image metadata.
func NewImagePartWithMeta(partName opc.PackURI, contentType string, blob []byte,
	pxWidth, pxHeight, horzDpi, vertDpi int, filename string,
) *ImagePart {
	return &ImagePart{
		BasePart:   opc.NewBasePart(partName, contentType, blob, nil),
		metaLoaded: true,
		pxWidth:    pxWidth,
		pxHeight:   pxHeight,
		horzDpi:    horzDpi,
		vertDpi:    vertDpi,
		filename:   filename,
	}
}

// NewImagePartFromImage creates an ImagePart from a parsed image.Image and its
// raw blob bytes. Used by the stream-based image insertion flow.
//
// Mirrors Python ImagePart.from_image(image, partname) classmethod.
func NewImagePartFromImage(img *image.Image, blob []byte) *ImagePart {
	fn := img.Filename()
	if fn == "" {
		fn = "image." + img.Ext()
	}
	return &ImagePart{
		BasePart:   opc.NewBasePart("", img.ContentType(), blob, nil),
		metaLoaded: true,
		pxWidth:    img.PxWidth(),
		pxHeight:   img.PxHeight(),
		horzDpi:    img.HorzDpi(),
		vertDpi:    img.VertDpi(),
		filename:   fn,
	}
}

// SHA1 returns the hex-encoded SHA1 hash of this image's blob.
// The value is cached after the first computation.
//
// Mirrors Python ImagePart.sha1 property.
func (ip *ImagePart) SHA1() string {
	if ip.sha1Hash != "" {
		return ip.sha1Hash
	}
	blob, _ := ip.Blob()
	h := sha1.Sum(blob)
	ip.sha1Hash = fmt.Sprintf("%x", h)
	return ip.sha1Hash
}

// Filename returns the original filename for this image. If no filename
// is available (nil/absent), a generic name based on the partname extension
// is returned.
//
// Mirrors Python ImagePart.filename property:
//
//	if self._image is not None:
//	    return self._image.filename
//	return "image.%s" % self.partname.ext
func (ip *ImagePart) Filename() string {
	if ip.filename != "" {
		return ip.filename
	}
	return "image." + ip.PartName().Ext()
}

// SetFilename sets the filename for this image part.
func (ip *ImagePart) SetFilename(fn string) {
	ip.filename = fn
}

// SetImageMeta sets the image dimensions and DPI metadata.
// Called by the image layer (MR-10) after parsing image headers.
func (ip *ImagePart) SetImageMeta(pxWidth, pxHeight, horzDpi, vertDpi int) {
	ip.pxWidth = pxWidth
	ip.pxHeight = pxHeight
	ip.horzDpi = horzDpi
	ip.vertDpi = vertDpi
	ip.metaLoaded = true
}

// ensureMeta lazily parses the image blob to populate dimensions and DPI
// metadata on first access. This mirrors Python's lazy Image property on
// ImagePart which parses the blob when first needed.
func (ip *ImagePart) ensureMeta() error {
	if ip.metaLoaded {
		return nil
	}
	blob, _ := ip.Blob()
	if len(blob) == 0 {
		return fmt.Errorf("parts: image part has no blob data")
	}
	img, err := image.FromBlob(blob, ip.Filename())
	if err != nil {
		return fmt.Errorf("parts: parsing image header: %w", err)
	}
	ip.pxWidth = img.PxWidth()
	ip.pxHeight = img.PxHeight()
	ip.horzDpi = img.HorzDpi()
	ip.vertDpi = img.VertDpi()
	if ip.filename == "" {
		ip.filename = img.Filename()
	}
	ip.metaLoaded = true
	return nil
}

// PxWidth returns the pixel width of this image.
func (ip *ImagePart) PxWidth() int {
	_ = ip.ensureMeta()
	return ip.pxWidth
}

// PxHeight returns the pixel height of this image.
func (ip *ImagePart) PxHeight() int {
	_ = ip.ensureMeta()
	return ip.pxHeight
}

// HorzDpi returns the horizontal dots per inch of this image.
func (ip *ImagePart) HorzDpi() int {
	_ = ip.ensureMeta()
	return ip.horzDpi
}

// VertDpi returns the vertical dots per inch of this image.
func (ip *ImagePart) VertDpi() int {
	_ = ip.ensureMeta()
	return ip.vertDpi
}

// --------------------------------------------------------------------------
// ImagePart.default_cx / default_cy — for embedded display size
// --------------------------------------------------------------------------

// DefaultCx returns the native width of this image in EMU for display.
// Calculated from pixel width and horizontal DPI using truncation.
//
// Mirrors Python ImagePart.default_cx:
//
//	Inches(px_width / horz_dpi) → int(px_width / horz_dpi * 914400)
func (ip *ImagePart) DefaultCx() (int64, error) {
	if err := ip.ensureMeta(); err != nil {
		return 0, err
	}
	if ip.horzDpi == 0 {
		return 0, fmt.Errorf("parts: image has no DPI metadata")
	}
	// Python: Inches(px_width / horz_dpi) → int(float * 914400) — TRUNCATES
	return int64(float64(ip.pxWidth) / float64(ip.horzDpi) * 914400), nil
}

// DefaultCy returns the native height of this image in EMU for display.
// NOTE: Python uses horz_dpi for cy (not vert_dpi) and round() (not int()).
//
// Mirrors Python ImagePart.default_cy:
//
//	Emu(int(round(914400 * px_height / horz_dpi)))
func (ip *ImagePart) DefaultCy() (int64, error) {
	if err := ip.ensureMeta(); err != nil {
		return 0, err
	}
	if ip.horzDpi == 0 {
		return 0, fmt.Errorf("parts: image has no DPI metadata")
	}
	// Python: int(round(914400 * px_height / horz_dpi)) — ROUNDS, uses horz_dpi
	return int64(math.Round(914400 * float64(ip.pxHeight) / float64(ip.horzDpi))), nil
}

// --------------------------------------------------------------------------
// Image.width / Image.height — for scaling computations
// --------------------------------------------------------------------------

// NativeWidth returns the native width of the image in EMU, matching
// Python Image.width.
//
// Mirrors Python Image.width:
//
//	Inches(self.px_width / self.horz_dpi) → int(px_width / horz_dpi * 914400)
func (ip *ImagePart) NativeWidth() (int64, error) {
	if err := ip.ensureMeta(); err != nil {
		return 0, err
	}
	if ip.horzDpi == 0 {
		return 0, fmt.Errorf("parts: image has no DPI metadata")
	}
	// Python: Inches() → int() truncation
	return int64(float64(ip.pxWidth) / float64(ip.horzDpi) * 914400), nil
}

// NativeHeight returns the native height of the image in EMU, matching
// Python Image.height. NOTE: uses vert_dpi (unlike DefaultCy which uses
// horz_dpi).
//
// Mirrors Python Image.height:
//
//	Inches(self.px_height / self.vert_dpi) → int(px_height / vert_dpi * 914400)
func (ip *ImagePart) NativeHeight() (int64, error) {
	if err := ip.ensureMeta(); err != nil {
		return 0, err
	}
	if ip.vertDpi == 0 {
		return 0, fmt.Errorf("parts: image has no vertical DPI metadata")
	}
	// Python: Inches() → int() truncation, uses VERT_dpi
	return int64(float64(ip.pxHeight) / float64(ip.vertDpi) * 914400), nil
}

// --------------------------------------------------------------------------
// ScaledDimensions — matches Python Image.scaled_dimensions EXACTLY
// --------------------------------------------------------------------------

// ScaledDimensions returns the scaled (cx, cy) in EMU for the given
// constraints. Uses NativeWidth/NativeHeight (which use vert_dpi for
// height), matching Python Image.scaled_dimensions.
//
// Mirrors Python Image.scaled_dimensions:
//
//	if width is None and height is None:
//	    return self.width, self.height
//	if width is None:
//	    scaling_factor = float(height) / float(self.height)
//	    width = round(self.width * scaling_factor)
//	if height is None:
//	    scaling_factor = float(width) / float(self.width)
//	    height = round(self.height * scaling_factor)
//	return Emu(width), Emu(height)
func (ip *ImagePart) ScaledDimensions(width, height *int64) (cx, cy int64, err error) {
	nativeW, err := ip.NativeWidth()
	if err != nil {
		return 0, 0, err
	}
	nativeH, err := ip.NativeHeight()
	if err != nil {
		return 0, 0, err
	}

	switch {
	case width == nil && height == nil:
		// Both nil → native size
		return nativeW, nativeH, nil
	case width == nil:
		// Width nil, height given → scale width from height
		if nativeH == 0 {
			return 0, *height, nil
		}
		scalingFactor := float64(*height) / float64(nativeH)
		w := int64(math.Round(float64(nativeW) * scalingFactor))
		return w, *height, nil
	case height == nil:
		// Height nil, width given → scale height from width
		if nativeW == 0 {
			return *width, 0, nil
		}
		scalingFactor := float64(*width) / float64(nativeW)
		h := int64(math.Round(float64(nativeH) * scalingFactor))
		return *width, h, nil
	default:
		// Both specified → use as-is
		return *width, *height, nil
	}
}

// --------------------------------------------------------------------------
// Factory / load
// --------------------------------------------------------------------------

// LoadImagePart is a PartConstructor that creates an ImagePart from package
// data during unmarshaling.
//
// Mirrors Python ImagePart.load classmethod:
//
//	return cls(partname, content_type, blob)  ← package ignored in Python
func LoadImagePart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	return NewImagePart(partName, contentType, blob, pkg), nil
}
