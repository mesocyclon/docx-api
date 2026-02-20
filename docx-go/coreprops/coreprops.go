// Package coreprops implements reading and writing of OPC core properties
// (docProps/core.xml) and extended/app properties (docProps/app.xml).
package coreprops

import (
	"encoding/xml"
	"strconv"
	"time"
)

// W3CDTF is the time format used by Dublin Core / XSD dateTime in OOXML.
const w3cdtf = "2006-01-02T15:04:05Z"

// CoreProperties represents the OPC core properties part (docProps/core.xml).
type CoreProperties struct {
	Title          string
	Subject        string
	Creator        string
	Keywords       string
	Description    string
	LastModifiedBy string
	Revision       string
	Created        time.Time
	Modified       time.Time
	Category       string
	ContentStatus  string
}

// AppProperties represents the extended properties part (docProps/app.xml).
type AppProperties struct {
	Template             string
	TotalTime            int
	Pages                int
	Words                int
	Characters           int
	Application          string
	DocSecurity          int
	Lines                int
	Paragraphs           int
	Company              string
	AppVersion           string
	CharactersWithSpaces int
}

// DefaultCore returns a CoreProperties with sensible defaults.
func DefaultCore(creator string) *CoreProperties {
	now := time.Now().UTC().Truncate(time.Second)
	return &CoreProperties{
		Creator:        creator,
		LastModifiedBy: creator,
		Revision:       "1",
		Created:        now,
		Modified:       now,
	}
}

// DefaultApp returns an AppProperties with sensible defaults.
func DefaultApp() *AppProperties {
	return &AppProperties{
		Template:    "Normal",
		Pages:       1,
		Lines:       1,
		Paragraphs:  1,
		Application: "docx-go",
		AppVersion:  "16.0000",
	}
}

// ---------------------------------------------------------------------------
// Core properties XML (de)serialization
// ---------------------------------------------------------------------------

// xmlCoreProperties is the internal XML mapping for core.xml.
type xmlCoreProperties struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties coreProperties"`

	Title          string     `xml:"http://purl.org/dc/elements/1.1/ title"`
	Subject        string     `xml:"http://purl.org/dc/elements/1.1/ subject"`
	Creator        string     `xml:"http://purl.org/dc/elements/1.1/ creator"`
	Keywords       string     `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties keywords"`
	Description    string     `xml:"http://purl.org/dc/elements/1.1/ description"`
	LastModifiedBy string     `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties lastModifiedBy"`
	Revision       string     `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties revision"`
	Created        *xmlW3CDTF `xml:"http://purl.org/dc/terms/ created"`
	Modified       *xmlW3CDTF `xml:"http://purl.org/dc/terms/ modified"`
	Category       string     `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties category,omitempty"`
	ContentStatus  string     `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties contentStatus,omitempty"`
}

// xmlW3CDTF represents a dcterms date element with xsi:type="dcterms:W3CDTF".
type xmlW3CDTF struct {
	Type  string `xml:"http://www.w3.org/2001/XMLSchema-instance type,attr,omitempty"`
	Value string `xml:",chardata"`
}

// ParseCore parses raw XML bytes from docProps/core.xml into CoreProperties.
func ParseCore(data []byte) (*CoreProperties, error) {
	var x xmlCoreProperties
	if err := xml.Unmarshal(data, &x); err != nil {
		return nil, err
	}

	cp := &CoreProperties{
		Title:          x.Title,
		Subject:        x.Subject,
		Creator:        x.Creator,
		Keywords:       x.Keywords,
		Description:    x.Description,
		LastModifiedBy: x.LastModifiedBy,
		Revision:       x.Revision,
		Category:       x.Category,
		ContentStatus:  x.ContentStatus,
	}

	if x.Created != nil && x.Created.Value != "" {
		if t, err := time.Parse(w3cdtf, x.Created.Value); err == nil {
			cp.Created = t
		}
	}
	if x.Modified != nil && x.Modified.Value != "" {
		if t, err := time.Parse(w3cdtf, x.Modified.Value); err == nil {
			cp.Modified = t
		}
	}

	return cp, nil
}

// SerializeCore produces the XML bytes for docProps/core.xml.
// It uses a custom encoder to emit correct namespace prefixes.
func SerializeCore(cp *CoreProperties) ([]byte, error) {
	x := xmlCoreProperties{
		Title:          cp.Title,
		Subject:        cp.Subject,
		Creator:        cp.Creator,
		Keywords:       cp.Keywords,
		Description:    cp.Description,
		LastModifiedBy: cp.LastModifiedBy,
		Revision:       cp.Revision,
		Category:       cp.Category,
		ContentStatus:  cp.ContentStatus,
	}
	if !cp.Created.IsZero() {
		x.Created = &xmlW3CDTF{
			Type:  "dcterms:W3CDTF",
			Value: cp.Created.UTC().Format(w3cdtf),
		}
	}
	if !cp.Modified.IsZero() {
		x.Modified = &xmlW3CDTF{
			Type:  "dcterms:W3CDTF",
			Value: cp.Modified.UTC().Format(w3cdtf),
		}
	}

	output, err := marshalCoreXML(&x)
	if err != nil {
		return nil, err
	}
	return append([]byte(xml.Header), output...), nil
}

// ---------------------------------------------------------------------------
// App properties XML (de)serialization
// ---------------------------------------------------------------------------

// xmlAppProperties is the internal XML mapping for app.xml.
type xmlAppProperties struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties Properties"`

	Template             string `xml:"Template,omitempty"`
	TotalTime            string `xml:"TotalTime,omitempty"`
	Pages                string `xml:"Pages,omitempty"`
	Words                string `xml:"Words,omitempty"`
	Characters           string `xml:"Characters,omitempty"`
	Application          string `xml:"Application,omitempty"`
	DocSecurity          string `xml:"DocSecurity,omitempty"`
	Lines                string `xml:"Lines,omitempty"`
	Paragraphs           string `xml:"Paragraphs,omitempty"`
	ScaleCrop            string `xml:"ScaleCrop,omitempty"`
	Company              string `xml:"Company"`
	LinksUpToDate        string `xml:"LinksUpToDate,omitempty"`
	CharactersWithSpaces string `xml:"CharactersWithSpaces,omitempty"`
	SharedDoc            string `xml:"SharedDoc,omitempty"`
	HyperlinksChanged    string `xml:"HyperlinksChanged,omitempty"`
	AppVersion           string `xml:"AppVersion,omitempty"`
}

// ParseApp parses raw XML bytes from docProps/app.xml into AppProperties.
func ParseApp(data []byte) (*AppProperties, error) {
	var x xmlAppProperties
	if err := xml.Unmarshal(data, &x); err != nil {
		return nil, err
	}

	return &AppProperties{
		Template:             x.Template,
		TotalTime:            atoi(x.TotalTime),
		Pages:                atoi(x.Pages),
		Words:                atoi(x.Words),
		Characters:           atoi(x.Characters),
		Application:          x.Application,
		DocSecurity:          atoi(x.DocSecurity),
		Lines:                atoi(x.Lines),
		Paragraphs:           atoi(x.Paragraphs),
		Company:              x.Company,
		AppVersion:           x.AppVersion,
		CharactersWithSpaces: atoi(x.CharactersWithSpaces),
	}, nil
}

// SerializeApp produces the XML bytes for docProps/app.xml.
func SerializeApp(ap *AppProperties) ([]byte, error) {
	x := xmlAppProperties{
		Template:             ap.Template,
		TotalTime:            itoa(ap.TotalTime),
		Pages:                itoa(ap.Pages),
		Words:                itoa(ap.Words),
		Characters:           itoa(ap.Characters),
		Application:          ap.Application,
		DocSecurity:          itoa(ap.DocSecurity),
		Lines:                itoa(ap.Lines),
		Paragraphs:           itoa(ap.Paragraphs),
		ScaleCrop:            "false",
		Company:              ap.Company,
		LinksUpToDate:        "false",
		CharactersWithSpaces: itoa(ap.CharactersWithSpaces),
		SharedDoc:            "false",
		HyperlinksChanged:    "false",
		AppVersion:           ap.AppVersion,
	}

	output, err := marshalAppXML(&x)
	if err != nil {
		return nil, err
	}
	return append([]byte(xml.Header), output...), nil
}

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

func atoi(s string) int {
	n, _ := strconv.Atoi(s)
	return n
}

func itoa(n int) string {
	return strconv.Itoa(n)
}
