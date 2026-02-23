package parts

import (
	"crypto/sha1"
	"fmt"
	"math"

	"github.com/vortex/go-docx/pkg/docx/opc"
)

// ImagePart stores an image as a binary blob. Unlike most parts, it does not
// contain XML — it embeds opc.BasePart directly.
//
// Mirrors Python ImagePart(Part).
type ImagePart struct {
	*opc.BasePart
	sha1Hash string // lazy, "" until first SHA1() call

	// Image metadata — populated after parsing. These fields are set when
	// the image layer (MR-10) is integrated. For now, ImageParts loaded
	// from packages carry blob data but metadata is deferred.
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
		BasePart: opc.NewBasePart(partName, contentType, blob, nil),
		pxWidth:  pxWidth,
		pxHeight: pxHeight,
		horzDpi:  horzDpi,
		vertDpi:  vertDpi,
		filename: filename,
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
// is available, a generic name based on the partname extension is returned.
//
// Mirrors Python ImagePart.filename property.
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
}

// PxWidth returns the pixel width of this image.
func (ip *ImagePart) PxWidth() int { return ip.pxWidth }

// PxHeight returns the pixel height of this image.
func (ip *ImagePart) PxHeight() int { return ip.pxHeight }

// HorzDpi returns the horizontal dots per inch of this image.
func (ip *ImagePart) HorzDpi() int { return ip.horzDpi }

// VertDpi returns the vertical dots per inch of this image.
func (ip *ImagePart) VertDpi() int { return ip.vertDpi }

// DefaultCx returns the native width of this image in EMU.
// Calculated from pixel width and horizontal DPI.
//
// Mirrors Python ImagePart.default_cx.
func (ip *ImagePart) DefaultCx() (int64, error) {
	if ip.horzDpi == 0 {
		return 0, fmt.Errorf("parts: image has no DPI metadata")
	}
	return int64(math.Round(float64(ip.pxWidth) / float64(ip.horzDpi) * 914400)), nil
}

// DefaultCy returns the native height of this image in EMU.
// NOTE: Python uses horz_dpi for cy too (not vert_dpi).
//
// Mirrors Python ImagePart.default_cy.
func (ip *ImagePart) DefaultCy() (int64, error) {
	if ip.horzDpi == 0 {
		return 0, fmt.Errorf("parts: image has no DPI metadata")
	}
	return int64(math.Round(914400 * float64(ip.pxHeight) / float64(ip.horzDpi))), nil
}

// Width returns the native width in EMU using horzDpi.
func (ip *ImagePart) Width() (int64, error) {
	return ip.DefaultCx()
}

// Height returns the native height in EMU using vertDpi.
func (ip *ImagePart) Height() (int64, error) {
	if ip.vertDpi == 0 {
		return 0, fmt.Errorf("parts: image has no DPI metadata")
	}
	return int64(math.Round(float64(ip.pxHeight) / float64(ip.vertDpi) * 914400)), nil
}

// ScaledDimensions returns the scaled (cx, cy) in EMU for the given
// constraints. If both width and height are nil, the native dimensions
// are returned. If only one is given, the other is scaled proportionally.
//
// Mirrors Python Image.scaled_dimensions EXACTLY.
func (ip *ImagePart) ScaledDimensions(width, height *int64) (cx, cy int64, err error) {
	nativeW, err := ip.DefaultCx()
	if err != nil {
		return 0, 0, err
	}
	nativeH, err := ip.DefaultCy()
	if err != nil {
		return 0, 0, err
	}

	switch {
	case width == nil && height == nil:
		// CASE 1: both nil → native size
		return nativeW, nativeH, nil
	case width == nil:
		// CASE 2: width nil → scale width from height
		if nativeH == 0 {
			return 0, *height, nil
		}
		scalingFactor := float64(*height) / float64(nativeH)
		w := int64(math.Round(float64(nativeW) * scalingFactor))
		return w, *height, nil
	case height == nil:
		// CASE 3: height nil → scale height from width
		if nativeW == 0 {
			return *width, 0, nil
		}
		scalingFactor := float64(*width) / float64(nativeW)
		h := int64(math.Round(float64(nativeH) * scalingFactor))
		return *width, h, nil
	default:
		// CASE 4: both specified → use as-is
		return *width, *height, nil
	}
}

// LoadImagePart is a PartConstructor that creates an ImagePart from package
// data during unmarshaling.
//
// Mirrors Python ImagePart.load classmethod.
func LoadImagePart(partName opc.PackURI, contentType, _ string, blob []byte, pkg *opc.OpcPackage) (opc.Part, error) {
	return NewImagePart(partName, contentType, blob, pkg), nil
}
