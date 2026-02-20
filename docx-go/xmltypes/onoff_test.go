package xmltypes

import (
	"encoding/xml"
	"testing"
)

func TestNewOnOff(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name     string
		input    bool
		wantVal  *string
		wantBool bool
	}{
		{"true creates nil Val", true, nil, true},
		{"false creates Val=0", false, strPtr("0"), false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := NewOnOff(tt.input)
			if tt.wantVal == nil && got.Val != nil {
				t.Errorf("expected nil Val, got %q", *got.Val)
			}
			if tt.wantVal != nil {
				if got.Val == nil {
					t.Errorf("expected Val=%q, got nil", *tt.wantVal)
				} else if *got.Val != *tt.wantVal {
					t.Errorf("expected Val=%q, got %q", *tt.wantVal, *got.Val)
				}
			}
			if got.Bool(false) != tt.wantBool {
				t.Errorf("Bool(false) = %v, want %v", got.Bool(false), tt.wantBool)
			}
		})
	}
}

func TestCTOnOff_Bool(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name       string
		onoff      *CT_OnOff
		defaultVal bool
		want       bool
	}{
		{"nil receiver returns default true", nil, true, true},
		{"nil receiver returns default false", nil, false, false},
		{"nil Val means true", &CT_OnOff{Val: nil}, false, true},
		{"Val=true", &CT_OnOff{Val: strPtr("true")}, false, true},
		{"Val=1", &CT_OnOff{Val: strPtr("1")}, false, true},
		{"Val=on", &CT_OnOff{Val: strPtr("on")}, false, true},
		{"Val=false", &CT_OnOff{Val: strPtr("false")}, true, false},
		{"Val=0", &CT_OnOff{Val: strPtr("0")}, true, false},
		{"Val=off", &CT_OnOff{Val: strPtr("off")}, true, false},
		{"unknown value defaults true", &CT_OnOff{Val: strPtr("maybe")}, false, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := tt.onoff.Bool(tt.defaultVal); got != tt.want {
				t.Errorf("Bool(%v) = %v, want %v", tt.defaultVal, got, tt.want)
			}
		})
	}
}

func TestCTOnOff_IsExplicitlySet(t *testing.T) {
	t.Parallel()

	if (*CT_OnOff)(nil).IsExplicitlySet() {
		t.Error("nil should not be explicitly set")
	}
	if !(&CT_OnOff{}).IsExplicitlySet() {
		t.Error("non-nil should be explicitly set")
	}
}

func TestCTOnOff_RoundTrip(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name string
		xml  string
	}{
		{
			"element without val",
			`<b xmlns="` + NSw + `"></b>`,
		},
		{
			"element with val=0",
			`<b xmlns="` + NSw + `"><val xmlns="` + NSw + `"></val></b>`,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var o1 CT_OnOff
			if err := xml.Unmarshal([]byte(tt.xml), &o1); err != nil {
				t.Fatalf("unmarshal: %v", err)
			}

			out, err := xml.Marshal(&o1)
			if err != nil {
				t.Fatalf("marshal: %v", err)
			}

			var o2 CT_OnOff
			if err := xml.Unmarshal(out, &o2); err != nil {
				t.Fatalf("re-unmarshal: %v", err)
			}

			if (o1.Val == nil) != (o2.Val == nil) {
				t.Error("round-trip changed Val nil-ness")
			}
			if o1.Val != nil && o2.Val != nil && *o1.Val != *o2.Val {
				t.Errorf("round-trip changed Val from %q to %q", *o1.Val, *o2.Val)
			}
		})
	}
}

func TestCTOnOff_UnmarshalWithAttr(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name    string
		xml     string
		wantNil bool
		wantVal string
	}{
		{
			"bare element",
			`<b xmlns="` + NSw + `"/>`,
			true, "",
		},
		{
			"val=false",
			`<b xmlns="` + NSw + `" val="false"/>`,
			false, "false",
		},
		{
			"val=0",
			`<b xmlns="` + NSw + `" val="0"/>`,
			false, "0",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var o CT_OnOff
			if err := xml.Unmarshal([]byte(tt.xml), &o); err != nil {
				t.Fatalf("unmarshal: %v", err)
			}
			if tt.wantNil && o.Val != nil {
				t.Errorf("expected nil Val, got %q", *o.Val)
			}
			if !tt.wantNil {
				if o.Val == nil {
					t.Error("expected non-nil Val")
				} else if *o.Val != tt.wantVal {
					t.Errorf("Val = %q, want %q", *o.Val, tt.wantVal)
				}
			}
		})
	}
}

func TestCTOnOff_MarshalTrue(t *testing.T) {
	t.Parallel()

	o := NewOnOff(true)
	out, err := xml.Marshal(o)
	if err != nil {
		t.Fatal(err)
	}
	// Should produce a self-closing element without val attribute
	s := string(out)
	if containsSubstring(s, "val=") {
		t.Errorf("true OnOff should not have val attribute, got: %s", s)
	}
}

func TestCTOnOff_MarshalFalse(t *testing.T) {
	t.Parallel()

	o := NewOnOff(false)
	out, err := xml.Marshal(o)
	if err != nil {
		t.Fatal(err)
	}
	s := string(out)
	if !containsSubstring(s, "val") {
		t.Errorf("false OnOff should have val attribute, got: %s", s)
	}
}

func containsSubstring(s, sub string) bool {
	return len(s) >= len(sub) && searchSubstring(s, sub)
}

func searchSubstring(s, sub string) bool {
	for i := 0; i <= len(s)-len(sub); i++ {
		if s[i:i+len(sub)] == sub {
			return true
		}
	}
	return false
}
