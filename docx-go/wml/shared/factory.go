package shared

import (
	"encoding/xml"
	"sync"
)

// ---------------------------------------------------------------------------
// BlockLevelFactory
// ---------------------------------------------------------------------------

// BlockLevelFactory creates a typed BlockLevelElement from an XML name.
// It is registered by downstream packages (body, para, table) via init().
// Returning nil means "I don't know this element" — the caller will then
// fall back to storing it as RawXML.
type BlockLevelFactory func(name xml.Name) BlockLevelElement

var (
	blockMu        sync.RWMutex
	blockFactories []BlockLevelFactory
)

// RegisterBlockFactory adds a factory to the block-level registry.
// It is safe to call concurrently (guarded by a mutex), but is typically
// called only from init() functions.
func RegisterBlockFactory(f BlockLevelFactory) {
	blockMu.Lock()
	defer blockMu.Unlock()
	blockFactories = append(blockFactories, f)
}

// CreateBlockElement asks every registered factory, in registration order,
// to create an element for the given XML name. The first non-nil result wins.
// If no factory recognises the name, nil is returned.
func CreateBlockElement(name xml.Name) BlockLevelElement {
	blockMu.RLock()
	defer blockMu.RUnlock()
	for _, f := range blockFactories {
		if el := f(name); el != nil {
			return el
		}
	}
	return nil
}

// ---------------------------------------------------------------------------
// ParagraphContentFactory
// ---------------------------------------------------------------------------

// ParagraphContentFactory creates a typed ParagraphContent from an XML name.
type ParagraphContentFactory func(name xml.Name) ParagraphContent

var (
	paraMu        sync.RWMutex
	paraFactories []ParagraphContentFactory
)

// RegisterParagraphContentFactory adds a factory to the paragraph-content
// registry.
func RegisterParagraphContentFactory(f ParagraphContentFactory) {
	paraMu.Lock()
	defer paraMu.Unlock()
	paraFactories = append(paraFactories, f)
}

// CreateParagraphContent asks every registered factory to create an element
// for the given XML name. Returns nil if unrecognised.
func CreateParagraphContent(name xml.Name) ParagraphContent {
	paraMu.RLock()
	defer paraMu.RUnlock()
	for _, f := range paraFactories {
		if el := f(name); el != nil {
			return el
		}
	}
	return nil
}

// ---------------------------------------------------------------------------
// RunContentFactory
// ---------------------------------------------------------------------------

// RunContentFactory creates a typed RunContent from an XML name.
type RunContentFactory func(name xml.Name) RunContent

var (
	runMu        sync.RWMutex
	runFactories []RunContentFactory
)

// RegisterRunContentFactory adds a factory to the run-content registry.
func RegisterRunContentFactory(f RunContentFactory) {
	runMu.Lock()
	defer runMu.Unlock()
	runFactories = append(runFactories, f)
}

// CreateRunContent asks every registered factory to create an element for
// the given XML name. Returns nil if unrecognised.
func CreateRunContent(name xml.Name) RunContent {
	runMu.RLock()
	defer runMu.RUnlock()
	for _, f := range runFactories {
		if el := f(name); el != nil {
			return el
		}
	}
	return nil
}
