// Package theme implements the word/theme/theme1.xml part as a raw XML passthrough.
//
// DrawingML theme schema is too complex for MVP parsing. Instead, theme bytes
// are stored as-is on open and written back verbatim on save.
//
// Contract: C-28 in contracts.md (0 dependencies — raw passthrough).
package theme

import "fmt"

// Parse accepts raw XML bytes of theme1.xml and returns them unchanged.
// An error is returned only if data is nil (defensive check).
func Parse(data []byte) ([]byte, error) {
	if data == nil {
		return nil, fmt.Errorf("theme: data is nil")
	}
	return data, nil
}

// Serialize accepts the stored theme bytes and returns them unchanged.
// An error is returned only if data is nil (defensive check).
func Serialize(data []byte) ([]byte, error) {
	if data == nil {
		return nil, fmt.Errorf("theme: data is nil")
	}
	return data, nil
}

// DefaultXML is the minimal valid theme1.xml that MS Word accepts without
// errors or repair prompts. It contains:
//   - clrScheme with 12 colours (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)
//   - fontScheme with majorFont and minorFont (latin, ea, cs)
//   - fmtScheme with fillStyleLst, lnStyleLst, effectStyleLst, bgFillStyleLst
//     (3 entries each — Word validates the count)
//
// Reference: reference-appendix.md §2.9, §5.12.
const DefaultXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" panose="020F0502020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`
