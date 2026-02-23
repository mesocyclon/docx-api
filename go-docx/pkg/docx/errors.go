// Package docx provides types and functions for creating and manipulating
// Office Open XML (.docx) documents.
package docx

import "fmt"

// DocxError is the base error type for all go-docx errors.
type DocxError struct {
	msg string
}

func (e *DocxError) Error() string { return e.msg }

// NewDocxError creates a new DocxError with the given message.
func NewDocxError(msg string, args ...any) *DocxError {
	return &DocxError{msg: fmt.Sprintf(msg, args...)}
}

// InvalidXmlError indicates that the XML is invalid or does not conform
// to the expected schema.
type InvalidXmlError struct {
	DocxError
}

// NewInvalidXmlError creates a new InvalidXmlError.
func NewInvalidXmlError(msg string, args ...any) *InvalidXmlError {
	return &InvalidXmlError{DocxError{msg: fmt.Sprintf(msg, args...)}}
}

// PackageNotFoundError indicates that a package file was not found.
type PackageNotFoundError struct {
	DocxError
}

// NewPackageNotFoundError creates a new PackageNotFoundError.
func NewPackageNotFoundError(msg string, args ...any) *PackageNotFoundError {
	return &PackageNotFoundError{DocxError{msg: fmt.Sprintf(msg, args...)}}
}

// InvalidSpanError indicates that a table cell span is invalid.
type InvalidSpanError struct {
	DocxError
}

// NewInvalidSpanError creates a new InvalidSpanError.
func NewInvalidSpanError(msg string, args ...any) *InvalidSpanError {
	return &InvalidSpanError{DocxError{msg: fmt.Sprintf(msg, args...)}}
}
