package opc

import (
	"fmt"
)

// --------------------------------------------------------------------------
// Serialized types (intermediate representations during reading)
// --------------------------------------------------------------------------

// SerializedRelationship is the intermediate representation of a relationship
// during package reading, before parts are resolved.
type SerializedRelationship struct {
	BaseURI    string
	RID        string
	RelType    string
	TargetRef  string
	TargetMode string // TargetModeInternal or TargetModeExternal
}

// IsExternal returns true if the relationship target is external.
func (sr SerializedRelationship) IsExternal() bool {
	return sr.TargetMode == TargetModeExternal
}

// TargetPartname resolves the target as a PackURI for internal relationships.
func (sr SerializedRelationship) TargetPartname() PackURI {
	return FromRelRef(sr.BaseURI, sr.TargetRef)
}

// SerializedPart holds the serialized data of a part read from the package.
type SerializedPart struct {
	Partname    PackURI
	ContentType string
	RelType     string
	Blob        []byte
	SRels       []SerializedRelationship
}

// --------------------------------------------------------------------------
// PackageReader
// --------------------------------------------------------------------------

// PackageReader reads an OPC package from a PhysPkgReader and produces
// serialized parts and relationships.
type PackageReader struct{}

// ReadResult holds the results of reading a package.
type ReadResult struct {
	PkgSRels []SerializedRelationship
	SParts   []SerializedPart
}

// Read reads the package and returns all serialized parts and relationships.
func (pr *PackageReader) Read(physReader *PhysPkgReader) (*ReadResult, error) {
	// 1. Parse [Content_Types].xml
	ctBlob, err := physReader.ContentTypesXml()
	if err != nil {
		return nil, fmt.Errorf("opc: reading content types: %w", err)
	}
	contentTypes, err := ParseContentTypes(ctBlob)
	if err != nil {
		return nil, err
	}

	// 2. Read package-level relationships
	pkgSRels, err := readSRels(physReader, PackageURI)
	if err != nil {
		return nil, fmt.Errorf("opc: reading package rels: %w", err)
	}

	// 3. Walk the relationship graph to discover all parts
	var sparts []SerializedPart
	visited := make(map[PackURI]bool)

	if err := walkParts(physReader, contentTypes, pkgSRels, &sparts, visited); err != nil {
		return nil, err
	}

	return &ReadResult{
		PkgSRels: pkgSRels,
		SParts:   sparts,
	}, nil
}

// walkParts recursively discovers parts by following relationships.
func walkParts(
	physReader *PhysPkgReader,
	contentTypes *ContentTypeMap,
	srels []SerializedRelationship,
	sparts *[]SerializedPart,
	visited map[PackURI]bool,
) error {
	for _, srel := range srels {
		if srel.IsExternal() {
			continue
		}
		partname := srel.TargetPartname()
		if visited[partname] {
			continue
		}
		visited[partname] = true

		blob, err := physReader.BlobFor(partname)
		if err != nil {
			return fmt.Errorf("opc: reading part %q: %w", partname, err)
		}

		ct, err := contentTypes.ContentType(partname)
		if err != nil {
			return err
		}

		partSRels, err := readSRels(physReader, partname)
		if err != nil {
			return fmt.Errorf("opc: reading rels for %q: %w", partname, err)
		}

		*sparts = append(*sparts, SerializedPart{
			Partname:    partname,
			ContentType: ct,
			RelType:     srel.RelType,
			Blob:        blob,
			SRels:       partSRels,
		})

		// Recurse into this part's relationships
		if err := walkParts(physReader, contentTypes, partSRels, sparts, visited); err != nil {
			return err
		}
	}
	return nil
}

// readSRels reads and parses the .rels file for the given source URI.
func readSRels(physReader *PhysPkgReader, sourceURI PackURI) ([]SerializedRelationship, error) {
	blob, err := physReader.RelsXmlFor(sourceURI)
	if err != nil {
		return nil, err
	}
	if blob == nil {
		return nil, nil
	}
	return ParseRelationships(blob, sourceURI.BaseURI())
}
