package codegen

import (
	"go/ast"
	"go/parser"
	"go/token"
	"strings"
	"testing"
)

// --- ExportName ---

func TestExportName(t *testing.T) {
	t.Parallel()
	tests := []struct {
		input, want string
	}{
		{"pPr", "PPr"},
		{"PPr", "PPr"},
		{"r", "R"},
		{"body", "Body"},
		{"", ""},
		{"Val", "Val"},
		{"rsidR", "RsidR"},
	}
	for _, tc := range tests {
		t.Run(tc.input, func(t *testing.T) {
			t.Parallel()
			if got := ExportName(tc.input); got != tc.want {
				t.Errorf("ExportName(%q) = %q, want %q", tc.input, got, tc.want)
			}
		})
	}
}

// --- Acceptance criterion: each cardinality generates correct number of methods ---

func TestZeroOrOne_Generates6Methods(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_PPr", Tag: "w:pPr"},
			{
				Name: "CT_P",
				Tag:  "w:p",
				Doc:  "paragraph element",
				Children: []Child{{
					Name: "PPr", Tag: "w:pPr", Type: "CT_PPr",
					Cardinality: "zero_or_one",
					Successors:  []string{"w:r", "w:hyperlink"},
				}},
			},
		},
	})

	// 6 methods: getter, GetOrAdd, Remove, add, new, insert
	assertContains(t, code, "func (e *CT_P) PPr() *CT_PPr")
	assertContains(t, code, "func (e *CT_P) GetOrAddPPr() *CT_PPr")
	assertContains(t, code, "func (e *CT_P) RemovePPr()")
	assertContains(t, code, "func (e *CT_P) addPPr() *CT_PPr")
	assertContains(t, code, "func (e *CT_P) newPPr() *CT_PPr")
	assertContains(t, code, "func (e *CT_P) insertPPr(child *CT_PPr) *CT_PPr")

	// Successors present in InsertElementBefore call
	assertContains(t, code, `"w:r"`)
	assertContains(t, code, `"w:hyperlink"`)
}

func TestZeroOrMore_Generates5Methods(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_R", Tag: "w:r"},
			{
				Name: "CT_P",
				Tag:  "w:p",
				Children: []Child{{
					Name: "R", Tag: "w:r", Type: "CT_R",
					Cardinality: "zero_or_more",
					Successors:  []string{"w:hyperlink"},
				}},
			},
		},
	})

	// 5 methods: list getter, Add (public), add, new, insert
	assertContains(t, code, "func (e *CT_P) RList() []*CT_R")
	assertContains(t, code, "func (e *CT_P) AddR() *CT_R")
	assertContains(t, code, "func (e *CT_P) addR() *CT_R")
	assertContains(t, code, "func (e *CT_P) newR() *CT_R")
	assertContains(t, code, "func (e *CT_P) insertR(child *CT_R) *CT_R")

	// Should NOT have singular getter (only list)
	assertNotContains(t, code, "func (e *CT_P) R() *CT_R")
}

func TestOneAndOnlyOne_Generates1Method(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_Body", Tag: "w:body"},
			{
				Name: "CT_Document",
				Tag:  "w:document",
				Children: []Child{{
					Name: "Body", Tag: "w:body", Type: "CT_Body",
					Cardinality: "one_and_only_one",
				}},
			},
		},
	})

	// 1 method: getter with panic
	assertContains(t, code, "func (e *CT_Document) Body() *CT_Body")
	assertContains(t, code, "panic(")

	// Should NOT have add/remove/insert methods
	assertNotContains(t, code, "func (e *CT_Document) AddBody()")
	assertNotContains(t, code, "func (e *CT_Document) RemoveBody()")
	assertNotContains(t, code, "func (e *CT_Document) GetOrAddBody()")
}

func TestOneOrMore_Generates5Methods(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_P", Tag: "w:p"},
			{
				Name: "CT_Tc",
				Tag:  "w:tc",
				Doc:  "table cell element",
				Children: []Child{{
					Name: "P", Tag: "w:p", Type: "CT_P",
					Cardinality: "one_or_more",
				}},
			},
		},
	})

	// 5 methods: list getter, Add (public), add, new, insert
	assertContains(t, code, "func (e *CT_Tc) PList() []*CT_P")
	assertContains(t, code, "func (e *CT_Tc) AddP() *CT_P")
	assertContains(t, code, "func (e *CT_Tc) addP() *CT_P")
	assertContains(t, code, "func (e *CT_Tc) newP() *CT_P")
	assertContains(t, code, "func (e *CT_Tc) insertP(child *CT_P) *CT_P")
	assertContains(t, code, "At least one must be present")
}

// --- Acceptance criterion: attributes ---

func TestOptionalAttribute_GeneratesGetterSetter(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{{
			Name: "CT_P",
			Tag:  "w:p",
			Attributes: []Attribute{{
				Name: "RsidR", AttrName: "w:rsidR",
				Type: "string", Required: false,
			}},
		}},
	})

	assertContains(t, code, "func (e *CT_P) RsidR() string")
	assertContains(t, code, "func (e *CT_P) SetRsidR(v string)")
	assertContains(t, code, "RemoveAttr")
}

func TestOptionalAttribute_IntType(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{{
			Name: "CT_Spacing",
			Tag:  "w:spacing",
			Attributes: []Attribute{{
				Name: "Before", AttrName: "w:before",
				Type: "int", Required: false,
			}},
		}},
	})

	assertContains(t, code, "func (e *CT_Spacing) Before() int")
	assertContains(t, code, "func (e *CT_Spacing) SetBefore(v int)")
	assertContains(t, code, "parseIntAttr(val)")
	assertContains(t, code, "formatIntAttr(v)")
}

func TestOptionalAttribute_BoolType(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{{
			Name: "CT_OnOff",
			Tag:  "w:onoff",
			Attributes: []Attribute{{
				Name: "Val", AttrName: "val",
				Type: "bool", Required: false,
			}},
		}},
	})

	assertContains(t, code, "func (e *CT_OnOff) Val() bool")
	assertContains(t, code, "parseBoolAttr(val)")
	assertContains(t, code, "formatBoolAttr(v)")
}

func TestRequiredAttribute_GeneratesGetterWithError(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{{
			Name: "CT_SchemeClr",
			Tag:  "a:schemeClr",
			Attributes: []Attribute{{
				Name: "Val", AttrName: "val",
				Type: "string", Required: true,
			}},
		}},
	})

	assertContains(t, code, "func (e *CT_SchemeClr) Val() (string, error)")
	assertContains(t, code, "func (e *CT_SchemeClr) SetVal(v string)")
	assertContains(t, code, "required attribute")
}

// --- Acceptance criterion: ZeroOrOneChoice ---

func TestChoiceGroup_GeneratesGroupAndPerChoiceMethods(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_SchemeClr", Tag: "a:schemeClr"},
			{Name: "CT_SrgbClr", Tag: "a:srgbClr"},
			{
				Name: "CT_Color",
				Tag:  "w:color",
				ChoiceGroups: []ChoiceGroup{{
					Name: "ColorChoice",
					Choices: []Choice{
						{Name: "SchemeClr", Tag: "a:schemeClr", Type: "CT_SchemeClr"},
						{Name: "SrgbClr", Tag: "a:srgbClr", Type: "CT_SrgbClr"},
					},
				}},
			},
		},
	})

	// Group getter and remover
	assertContains(t, code, "func (e *CT_Color) ColorChoice() *Element")
	assertContains(t, code, "func (e *CT_Color) RemoveColorChoice()")

	// Per-choice methods (5 each: getter, GetOrChangeTo, add, new, insert)
	assertContains(t, code, "func (e *CT_Color) SchemeClr() *CT_SchemeClr")
	assertContains(t, code, "func (e *CT_Color) GetOrChangeToSchemeClr() *CT_SchemeClr")
	assertContains(t, code, "func (e *CT_Color) addSchemeClr() *CT_SchemeClr")
	assertContains(t, code, "func (e *CT_Color) newSchemeClr() *CT_SchemeClr")
	assertContains(t, code, "func (e *CT_Color) insertSchemeClr(child *CT_SchemeClr) *CT_SchemeClr")

	assertContains(t, code, "func (e *CT_Color) SrgbClr() *CT_SrgbClr")
	assertContains(t, code, "func (e *CT_Color) GetOrChangeToSrgbClr() *CT_SrgbClr")

	// GetOrChangeTo should call RemoveColorChoice
	assertContains(t, code, "e.RemoveColorChoice()")
}

// --- Acceptance criterion: generated code is valid Go ---

func TestGenerate_OutputIsParsableGo(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_PPr", Tag: "w:pPr"},
			{Name: "CT_R", Tag: "w:r"},
			{Name: "CT_Hyperlink", Tag: "w:hyperlink"},
			{Name: "CT_Body", Tag: "w:body"},
			{
				Name: "CT_P",
				Tag:  "w:p",
				Doc:  "paragraph element",
				Children: []Child{
					{Name: "PPr", Tag: "w:pPr", Type: "CT_PPr", Cardinality: "zero_or_one", Successors: []string{"w:r", "w:hyperlink"}},
					{Name: "R", Tag: "w:r", Type: "CT_R", Cardinality: "zero_or_more", Successors: []string{"w:hyperlink"}},
					{Name: "Hyperlink", Tag: "w:hyperlink", Type: "CT_Hyperlink", Cardinality: "zero_or_more"},
				},
				Attributes: []Attribute{
					{Name: "RsidR", AttrName: "w:rsidR", Type: "string"},
				},
			},
			{
				Name: "CT_Document",
				Tag:  "w:document",
				Children: []Child{
					{Name: "Body", Tag: "w:body", Type: "CT_Body", Cardinality: "one_and_only_one"},
				},
			},
			{
				Name: "CT_Tc",
				Tag:  "w:tc",
				Children: []Child{
					{Name: "P", Tag: "w:p", Type: "CT_P", Cardinality: "one_or_more"},
				},
			},
		},
	})

	// Parse as valid Go
	fset := token.NewFileSet()
	f, err := parser.ParseFile(fset, "generated.go", code, parser.AllErrors)
	if err != nil {
		t.Fatalf("generated code is not valid Go:\n%v\n--- code ---\n%s", err, code)
	}

	// Verify package name
	if f.Name.Name != "oxml" {
		t.Errorf("package name = %q, want %q", f.Name.Name, "oxml")
	}

	// Verify struct declarations exist
	structs := collectStructNames(f)
	for _, name := range []string{"CT_P", "CT_PPr", "CT_R", "CT_Hyperlink", "CT_Body", "CT_Document", "CT_Tc"} {
		if !structs[name] {
			t.Errorf("expected struct %q not found in generated code", name)
		}
	}

	// Verify method count for CT_P
	methods := collectMethodNames(f, "CT_P")
	// zero_or_one PPr: 6 methods
	for _, m := range []string{"PPr", "GetOrAddPPr", "RemovePPr", "addPPr", "newPPr", "insertPPr"} {
		if !methods[m] {
			t.Errorf("expected method CT_P.%s not found", m)
		}
	}
	// zero_or_more R: 5 methods
	for _, m := range []string{"RList", "AddR", "addR", "newR", "insertR"} {
		if !methods[m] {
			t.Errorf("expected method CT_P.%s not found", m)
		}
	}
	// optional attribute RsidR: 2 methods
	for _, m := range []string{"RsidR", "SetRsidR"} {
		if !methods[m] {
			t.Errorf("expected method CT_P.%s not found", m)
		}
	}
}

// --- Structure tests ---

func TestGenerate_StructEmbedsElement(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package:  "oxml",
		Elements: []Element{{Name: "CT_P", Tag: "w:p"}},
	})

	assertContains(t, code, "type CT_P struct")
	assertContains(t, code, "Element")
}

func TestGenerate_HeaderComment(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package:  "oxml",
		Elements: []Element{{Name: "CT_P", Tag: "w:p"}},
	})

	if !strings.HasPrefix(code, "// Code generated") {
		t.Errorf("output should start with generated file header, got: %s", code[:80])
	}
	assertContains(t, code, "DO NOT EDIT")
}

func TestGenerate_DocComment(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package:  "oxml",
		Elements: []Element{{Name: "CT_P", Tag: "w:p", Doc: "paragraph element"}},
	})

	assertContains(t, code, "paragraph element")
}

func TestGenerate_EmptySchema(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package:  "oxml",
		Elements: []Element{},
	})

	assertContains(t, code, "package oxml")
}

func TestGenerate_MultipleElements(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_A", Tag: "w:a"},
			{Name: "CT_B", Tag: "w:b"},
			{Name: "CT_C", Tag: "w:c"},
		},
	})

	assertContains(t, code, "type CT_A struct")
	assertContains(t, code, "type CT_B struct")
	assertContains(t, code, "type CT_C struct")
}

func TestGenerate_NoSuccessors_CallsInsertWithoutArgs(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_Child", Tag: "w:child"},
			{
				Name: "CT_Parent",
				Tag:  "w:parent",
				Children: []Child{{
					Name: "Child", Tag: "w:child", Type: "CT_Child",
					Cardinality: "zero_or_one",
					Successors:  []string{},
				}},
			},
		},
	})

	// InsertElementBefore with no successors → just child.E, no extra args
	assertContains(t, code, "e.InsertElementBefore(child.E)")
}

func TestGenerate_SuccessorsPreserveOrder(t *testing.T) {
	t.Parallel()
	code := generateCode(t, Schema{
		Package: "oxml",
		Elements: []Element{
			{Name: "CT_A", Tag: "w:a"},
			{
				Name: "CT_Parent",
				Tag:  "w:parent",
				Children: []Child{{
					Name: "A", Tag: "w:a", Type: "CT_A",
					Cardinality: "zero_or_one",
					Successors:  []string{"w:b", "w:c", "w:d"},
				}},
			},
		},
	})

	// All three successors should appear in order
	assertContains(t, code, `"w:b", "w:c", "w:d"`)
}

// --- AttrType resolution ---

func TestResolveAttrType_String(t *testing.T) {
	t.Parallel()
	goType, zero, def, parse, format := resolveAttrType(Attribute{Type: "string"})
	assertEqual(t, "string", goType)
	assertEqual(t, `""`, zero)
	assertEqual(t, `""`, def)
	assertEqual(t, "val", parse)
	assertEqual(t, "v", format)
}

func TestResolveAttrType_Int(t *testing.T) {
	t.Parallel()
	goType, _, _, parse, format := resolveAttrType(Attribute{Type: "int"})
	assertEqual(t, "int", goType)
	assertEqual(t, "parseIntAttr(val)", parse)
	assertEqual(t, "formatIntAttr(v)", format)
}

func TestResolveAttrType_Bool(t *testing.T) {
	t.Parallel()
	goType, _, _, parse, format := resolveAttrType(Attribute{Type: "bool"})
	assertEqual(t, "bool", goType)
	assertEqual(t, "parseBoolAttr(val)", parse)
	assertEqual(t, "formatBoolAttr(v)", format)
}

func TestResolveAttrType_Int64(t *testing.T) {
	t.Parallel()
	goType, _, _, parse, format := resolveAttrType(Attribute{Type: "int64"})
	assertEqual(t, "int64", goType)
	assertEqual(t, "parseInt64Attr(val)", parse)
	assertEqual(t, "formatInt64Attr(v)", format)
}

// --- Helpers ---

func generateCode(t *testing.T, schema Schema) string {
	t.Helper()
	gen, err := NewGenerator(schema)
	if err != nil {
		t.Fatalf("NewGenerator error: %v", err)
	}
	output, err := gen.Generate()
	if err != nil {
		t.Fatalf("Generate error: %v", err)
	}
	return string(output)
}

func assertContains(t *testing.T, s, substr string) {
	t.Helper()
	if !strings.Contains(s, substr) {
		t.Errorf("output does not contain %q\n--- output (first 2000 chars) ---\n%s", substr, truncate(s, 2000))
	}
}

func assertNotContains(t *testing.T, s, substr string) {
	t.Helper()
	if strings.Contains(s, substr) {
		t.Errorf("output should NOT contain %q", substr)
	}
}

func assertEqual(t *testing.T, want, got string) {
	t.Helper()
	if got != want {
		t.Errorf("got %q, want %q", got, want)
	}
}

func truncate(s string, n int) string {
	if len(s) <= n {
		return s
	}
	return s[:n] + "..."
}

// collectStructNames returns a set of type names declared as structs in the AST.
func collectStructNames(f *ast.File) map[string]bool {
	names := make(map[string]bool)
	for _, decl := range f.Decls {
		gd, ok := decl.(*ast.GenDecl)
		if !ok || gd.Tok != token.TYPE {
			continue
		}
		for _, spec := range gd.Specs {
			ts, ok := spec.(*ast.TypeSpec)
			if !ok {
				continue
			}
			if _, ok := ts.Type.(*ast.StructType); ok {
				names[ts.Name.Name] = true
			}
		}
	}
	return names
}

// collectMethodNames returns a set of method names for the given receiver type.
func collectMethodNames(f *ast.File, receiverType string) map[string]bool {
	names := make(map[string]bool)
	for _, decl := range f.Decls {
		fd, ok := decl.(*ast.FuncDecl)
		if !ok || fd.Recv == nil || len(fd.Recv.List) == 0 {
			continue
		}

		// Extract receiver type name
		recvType := ""
		switch rt := fd.Recv.List[0].Type.(type) {
		case *ast.StarExpr:
			if ident, ok := rt.X.(*ast.Ident); ok {
				recvType = ident.Name
			}
		case *ast.Ident:
			recvType = rt.Name
		}

		if recvType == receiverType {
			names[fd.Name.Name] = true
		}
	}
	return names
}
