package opc

import (
	"encoding/xml"

	"github.com/vortex/docx-go/xmltypes"
)

// xmlRelationships represents a .rels XML document.
type xmlRelationships struct {
	XMLName       xml.Name          `xml:"Relationships"`
	Relationships []xmlRelationship `xml:"Relationship"`
}

type xmlRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// parseRels parses a .rels XML document into a slice of Relationship.
func parseRels(data []byte) ([]Relationship, error) {
	var xr xmlRelationships
	if err := xml.Unmarshal(data, &xr); err != nil {
		return nil, err
	}
	rels := make([]Relationship, len(xr.Relationships))
	for i, r := range xr.Relationships {
		rels[i] = Relationship(r)
	}
	return rels, nil
}

// buildRels serializes relationships into .rels XML bytes.
func buildRels(rels []Relationship) ([]byte, error) {
	if len(rels) == 0 {
		return nil, nil
	}
	xr := xmlRelationships{
		XMLName: xml.Name{Space: xmltypes.NSRelationships, Local: "Relationships"},
	}
	for _, r := range rels {
		xr.Relationships = append(xr.Relationships, xmlRelationship(r))
	}
	return marshalXMLWithHeader(xr)
}
