// Package websettings implements raw passthrough for word/webSettings.xml.
//
// The webSettings part is mandatory in a .docx package (even if empty),
// but its content does not need to be parsed for the MVP.
// Parse stores the raw bytes; Serialize returns them unchanged.
package websettings

// Parse accepts raw XML bytes from word/webSettings.xml and returns them as-is.
// The data is stored opaquely and round-tripped without modification.
func Parse(data []byte) ([]byte, error) {
	return data, nil
}

// Serialize returns the raw XML bytes to be written back to word/webSettings.xml.
func Serialize(data []byte) ([]byte, error) {
	return data, nil
}
