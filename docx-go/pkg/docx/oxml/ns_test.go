package oxml

import (
	"strings"
	"testing"
)

func TestQn(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name string
		tag  string
		want string
	}{
		{
			name: "w:p resolves to wordprocessingml namespace",
			tag:  "w:p",
			want: "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p",
		},
		{
			name: "w:body resolves correctly",
			tag:  "w:body",
			want: "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body",
		},
		{
			name: "r:id resolves to relationships namespace",
			tag:  "r:id",
			want: "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
		},
		{
			name: "a:blip resolves to drawingml namespace",
			tag:  "a:blip",
			want: "{http://schemas.openxmlformats.org/drawingml/2006/main}blip",
		},
		{
			name: "wp:inline resolves to wordprocessingDrawing namespace",
			tag:  "wp:inline",
			want: "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline",
		},
		{
			name: "pic:pic resolves to picture namespace",
			tag:  "pic:pic",
			want: "{http://schemas.openxmlformats.org/drawingml/2006/picture}pic",
		},
		{
			name: "w14:conflictMode resolves to word 2010 namespace",
			tag:  "w14:conflictMode",
			want: "{http://schemas.microsoft.com/office/word/2010/wordml}conflictMode",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			got := Qn(tt.tag)
			if got != tt.want {
				t.Errorf("Qn(%q) = %q, want %q", tt.tag, got, tt.want)
			}
		})
	}
}

func TestQnPanicsOnUnknownPrefix(t *testing.T) {
	t.Parallel()
	defer func() {
		if r := recover(); r == nil {
			t.Error("expected panic for unknown prefix, got nil")
		}
	}()
	Qn("unknown:tag")
}

func TestQnNoPrefix(t *testing.T) {
	t.Parallel()
	got := Qn("simpleTag")
	if got != "simpleTag" {
		t.Errorf("Qn(%q) = %q, want %q", "simpleTag", got, "simpleTag")
	}
}

func TestNewNSPTag(t *testing.T) {
	t.Parallel()
	tag := NewNSPTag("w:p")
	if tag.Prefix() != "w" {
		t.Errorf("Prefix() = %q, want %q", tag.Prefix(), "w")
	}
	if tag.LocalPart() != "p" {
		t.Errorf("LocalPart() = %q, want %q", tag.LocalPart(), "p")
	}
	if tag.NsURI() != Nsmap["w"] {
		t.Errorf("NsURI() = %q, want %q", tag.NsURI(), Nsmap["w"])
	}
	if tag.String() != "w:p" {
		t.Errorf("String() = %q, want %q", tag.String(), "w:p")
	}
}

func TestNewNSPTagClarkName(t *testing.T) {
	t.Parallel()
	tag := NewNSPTag("w:body")
	want := Qn("w:body")
	if tag.ClarkName() != want {
		t.Errorf("ClarkName() = %q, want %q", tag.ClarkName(), want)
	}
}

func TestNSPTagFromClark(t *testing.T) {
	t.Parallel()
	clark := Qn("w:p")
	tag := NSPTagFromClark(clark)
	if tag.String() != "w:p" {
		t.Errorf("String() = %q, want %q", tag.String(), "w:p")
	}
	if tag.ClarkName() != clark {
		t.Errorf("ClarkName() = %q, want %q", tag.ClarkName(), clark)
	}
}

func TestNSPTagRoundTrip(t *testing.T) {
	t.Parallel()
	tags := []string{"w:p", "w:body", "r:id", "a:blip", "wp:inline", "pic:pic", "w14:conflictMode"}
	for _, nstag := range tags {
		t.Run(nstag, func(t *testing.T) {
			t.Parallel()
			clark := NewNSPTag(nstag).ClarkName()
			back := NSPTagFromClark(clark)
			if back.String() != nstag {
				t.Errorf("round-trip failed: %q → %q → %q", nstag, clark, back.String())
			}
		})
	}
}

func TestNsDecls(t *testing.T) {
	t.Parallel()
	t.Run("single prefix", func(t *testing.T) {
		t.Parallel()
		decl := NsDecls("w")
		if !strings.Contains(decl, `xmlns:w=`) {
			t.Errorf("NsDecls('w') = %q, expected to contain xmlns:w=", decl)
		}
		if !strings.Contains(decl, Nsmap["w"]) {
			t.Errorf("NsDecls('w') = %q, expected to contain URI", decl)
		}
	})

	t.Run("multiple prefixes", func(t *testing.T) {
		t.Parallel()
		decl := NsDecls("w", "r")
		if !strings.Contains(decl, `xmlns:w=`) {
			t.Errorf("NsDecls('w', 'r') missing xmlns:w")
		}
		if !strings.Contains(decl, `xmlns:r=`) {
			t.Errorf("NsDecls('w', 'r') missing xmlns:r")
		}
	})
}

func TestNsPfxMap(t *testing.T) {
	t.Parallel()
	m := NsPfxMap("w", "r")
	if len(m) != 2 {
		t.Fatalf("NsPfxMap returned %d entries, want 2", len(m))
	}
	if m["w"] != Nsmap["w"] {
		t.Errorf("m[w] = %q, want %q", m["w"], Nsmap["w"])
	}
	if m["r"] != Nsmap["r"] {
		t.Errorf("m[r] = %q, want %q", m["r"], Nsmap["r"])
	}
}

func TestNSPTagNsMap(t *testing.T) {
	t.Parallel()
	tag := NewNSPTag("w:p")
	m := tag.NsMap()
	if len(m) != 1 {
		t.Fatalf("NsMap returned %d entries, want 1", len(m))
	}
	if m["w"] != Nsmap["w"] {
		t.Errorf("m[w] = %q, want %q", m["w"], Nsmap["w"])
	}
}

func TestPfxmapIsInverseOfNsmap(t *testing.T) {
	t.Parallel()
	for pfx, uri := range Nsmap {
		got, ok := Pfxmap[uri]
		if !ok {
			t.Errorf("Pfxmap missing URI %q (prefix %q)", uri, pfx)
			continue
		}
		if got != pfx {
			t.Errorf("Pfxmap[%q] = %q, want %q", uri, got, pfx)
		}
	}
}
