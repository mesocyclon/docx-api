package packaging

import (
	"fmt"
	"path"
	"strings"
)

// NextRelID generates a unique relationship ID (e.g. "rId7") that does not
// collide with any existing relationship. It is safe for concurrent use.
func (d *Document) NextRelID() string {
	d.mu.Lock()
	defer d.mu.Unlock()
	id := fmt.Sprintf("rId%d", d.nextRelSeq)
	d.nextRelSeq++
	return id
}

// NextBookmarkID returns a unique numeric ID suitable for w:bookmarkStart,
// w:commentRangeStart, w:ins, w:del, and similar elements that require a
// document-wide unique integer ID. Safe for concurrent use.
func (d *Document) NextBookmarkID() int {
	d.mu.Lock()
	defer d.mu.Unlock()
	id := d.nextBmkID
	d.nextBmkID++
	return id
}

// AddMedia embeds a media file (typically an image) into the document.
// The filename should include its extension (e.g. "photo.png").
// The returned string is the rId that can be referenced from a drawing
// element's r:embed attribute.
//
// The caller is responsible for using the returned rId in the document XML
// (e.g. inside a wp:inline / a:blip).
func (d *Document) AddMedia(filename string, data []byte) string {
	d.mu.Lock()
	defer d.mu.Unlock()

	// Ensure unique filename in the media map.
	base := strings.TrimSuffix(filename, path.Ext(filename))
	ext := path.Ext(filename)
	final := filename
	for i := 1; ; i++ {
		if _, exists := d.Media[final]; !exists {
			break
		}
		final = fmt.Sprintf("%s%d%s", base, i, ext)
	}

	d.Media[final] = data

	// Generate a relationship ID. The actual relationship will be created
	// during buildPackage, but we return the rId now so it can be wired
	// into the XML tree.
	rID := fmt.Sprintf("rId%d", d.nextRelSeq)
	d.nextRelSeq++

	return rID
}
