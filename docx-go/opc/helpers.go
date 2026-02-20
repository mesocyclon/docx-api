package opc

import (
	"fmt"
	"strconv"
	"strings"
)

// normalizeName ensures part names start with "/" and use forward slashes.
func normalizeName(name string) string {
	name = strings.ReplaceAll(name, "\\", "/")
	if !strings.HasPrefix(name, "/") {
		name = "/" + name
	}
	return name
}

// parseRelID extracts the numeric part from "rIdN". Returns 0 if unparseable.
func parseRelID(id string) int {
	if strings.HasPrefix(id, "rId") {
		n, err := strconv.Atoi(id[3:])
		if err == nil {
			return n
		}
	}
	return 0
}

// formatRelID formats a number as "rIdN".
func formatRelID(n int) string {
	return fmt.Sprintf("rId%d", n)
}

// nextRelIDFrom finds the max rId in a slice of relationships and returns the next one.
func nextRelIDFrom(rels []Relationship) string {
	max := 0
	for _, r := range rels {
		if n := parseRelID(r.ID); n > max {
			max = n
		}
	}
	return formatRelID(max + 1)
}

// relsZipPath returns the .rels ZIP entry path for a given part name.
// e.g., "/word/document.xml" → "word/_rels/document.xml.rels"
// Package-level rels: "" or "/" → "_rels/.rels"
func relsZipPath(partName string) string {
	if partName == "" || partName == "/" {
		return "_rels/.rels"
	}
	partName = normalizeName(partName)
	idx := strings.LastIndex(partName, "/")
	dir := partName[1 : idx+1] // strip leading "/", keep trailing "/"
	base := partName[idx+1:]
	return dir + "_rels/" + base + ".rels"
}

// stripLeadingSlash removes the leading "/" for ZIP storage.
func stripLeadingSlash(name string) string {
	return strings.TrimPrefix(name, "/")
}
