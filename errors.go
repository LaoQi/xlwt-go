package xlwt

import "errors"

// Limits of the BIFF8 file format. Values beyond these ranges cannot be
// represented in a .xls file; they are rejected instead of silently producing a
// corrupt workbook.
const (
	// MaxRow is the highest zero-based row index (BIFF8 supports 65536 rows).
	MaxRow = 65535
	// MaxCol is the highest zero-based column index (BIFF8 supports 256 columns).
	MaxCol = 255
	// MaxStringLength is the longest string that fits into a single BIFF8 string
	// cell, measured in characters.
	MaxStringLength = 32767
	// MaxSheetNameLength is the longest allowed worksheet name, in characters.
	MaxSheetNameLength = 31
)

// Errors returned by Write and AddSheet.
var (
	// ErrRowOutOfRange is returned when a row index falls outside [0, MaxRow].
	ErrRowOutOfRange = errors.New("xlwt: row index out of range")
	// ErrColOutOfRange is returned when a column index falls outside [0, MaxCol].
	ErrColOutOfRange = errors.New("xlwt: column index out of range")
	// ErrStringTooLong is returned when a string exceeds MaxStringLength characters.
	ErrStringTooLong = errors.New("xlwt: string exceeds the BIFF8 limit")
	// ErrInvalidSheetName is returned for an empty name, a name longer than
	// MaxSheetNameLength characters, or a name containing characters that Excel
	// does not allow: \ / ? * [ ] : and a leading apostrophe.
	ErrInvalidSheetName = errors.New("xlwt: invalid worksheet name")
	// ErrInvalidColumnWidth is returned when a column width cannot be
	// represented in a COLINFO record.
	ErrInvalidColumnWidth = errors.New("xlwt: invalid column width")
	// ErrInvalidRowHeight is returned when a row height cannot be represented in
	// a ROW record.
	ErrInvalidRowHeight = errors.New("xlwt: invalid row height")
	// ErrUnattachedWorksheet is returned when an explicit style is used on a
	// worksheet that was created with NewWorksheet instead of Workbook.AddSheet.
	ErrUnattachedWorksheet = errors.New("xlwt: worksheet is not attached to a workbook")
	// ErrTooManyStyles is returned when more distinct cell styles are used than a
	// BIFF8 file can reference (4094).
	ErrTooManyStyles = errors.New("xlwt: too many distinct cell styles")
	// ErrTooManyNumberFormats is returned when more user defined number formats
	// are used than a BIFF8 file can reference (starting at index 164).
	ErrTooManyNumberFormats = errors.New("xlwt: too many custom number formats")
	// ErrDuplicateSheetName is returned when a worksheet name is already used by
	// another sheet of the same workbook (comparison is case-insensitive).
	ErrDuplicateSheetName = errors.New("xlwt: duplicate worksheet name")
)
