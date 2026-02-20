# wml/table — Module C-13 Implementation

## File Structure

```
pkg/
├── go.mod                          ← module root
├── xmltypes/xmltypes.go            ← dependency stub (C-02)
├── wml/shared/shared.go            ← dependency stub (C-05)
└── wml/table/                      ← MODULE IMPLEMENTATION
    ├── types.go                    ← All type definitions (CT_Tbl, CT_Row, CT_Tc, etc.)
    ├── helpers.go                  ← Encoding/decoding utilities (encodeChild, encodeRawSlice)
    ├── util.go                     ← String/int conversions
    ├── marshal_tbl.go              ← MarshalXML/UnmarshalXML for CT_Tbl, CT_TblGrid
    ├── marshal_tblpr.go            ← MarshalXML/UnmarshalXML for CT_TblPr, CT_TblBorders, CT_TblCellMar
    ├── marshal_row.go              ← MarshalXML/UnmarshalXML for CT_Row
    ├── marshal_tc.go               ← MarshalXML/UnmarshalXML for CT_Tc, CT_TcPr, CT_TcBorders
    ├── marshal_trpr.go             ← MarshalXML/UnmarshalXML for CT_TrPr, CT_TblPrEx
    └── table_test.go               ← Round-trip and structural tests
```

## Key Design Decisions

1. **Strict XSD sequence order** — CT_TblPr, CT_TcPr, CT_TrPr use hand-coded MarshalXML (no reflect) to guarantee element ordering per the WML XSD specification.

2. **RawXML round-trip** — Unknown/extension elements (e.g. `w14:*`) are captured as `shared.RawXML` in `Extra` slices and faithfully re-serialized during marshal.

3. **Interface wrappers** — `RawTblContent` and `RawRowContent` wrap `shared.RawXML` to satisfy the package-local `TblContent`/`RowContent` interfaces without creating import cycles.

4. **Block factory delegation** — `CT_Tc.UnmarshalXML` uses `shared.CreateBlockElement()` to resolve cross-package types (e.g. paragraphs registered by the `para` module via `init()`). Unknown elements fall back to `shared.RawXML`.

5. **Transitional + left/right aliases** — Border and cell-margin unmarshalers accept both `start`/`end` (Strict) and `left`/`right` (Transitional) element names.

6. **Nested tables** — `CT_Tc` directly handles `<w:tbl>` children for recursive table support.

## Dependencies

- `xmltypes` — namespace constants, CT_OnOff, CT_String, CT_DecimalNumber, CT_Border, CT_Shd
- `wml/shared` — BlockLevelElement interface, RawXML, block factory

## Tests

Tests cover:
- Full 2×2 table round-trip (from reference-appendix section 3.2)
- TblPr XSD element ordering verification
- TcPr XSD element ordering verification
- TrPr round-trip
- RawXML extension element preservation
- Table borders round-trip
- Cell margins round-trip
- Nested table parsing
- VMerge with/without val attribute
