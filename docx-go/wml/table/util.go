package table

import (
	"bytes"
	"io"
	"strconv"
)

// intToStr converts int to string for XML attribute values.
func intToStr(v int) string {
	return strconv.Itoa(v)
}

// strToInt converts string to int; returns 0 on error.
func strToInt(s string) int {
	v, _ := strconv.Atoi(s)
	return v
}

// bytesReader wraps a byte slice as an io.Reader.
func bytesReader(b []byte) io.Reader {
	return bytes.NewReader(b)
}
