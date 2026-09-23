# xlwt-go

[![CI](https://github.com/LaoQi/xlwt-go/actions/workflows/ci.yml/badge.svg)](https://github.com/LaoQi/xlwt-go/actions/workflows/ci.yml)
[![Go Reference](https://pkg.go.dev/badge/github.com/LaoQi/xlwt-go.svg)](https://pkg.go.dev/github.com/LaoQi/xlwt-go)
[![License: LGPL-2.1](https://img.shields.io/badge/License-LGPL--2.1-blue.svg)](LICENSE)

A pure Go library for writing Microsoft Excel 97-2003 (`.xls`, BIFF8) files.
Ported from the Python [xlwt](https://github.com/python-excel/xlwt) library.
No third-party dependencies, standard library only.

The generated files can be opened by Excel, WPS, LibreOffice and read by tools
such as Python's [xlrd](https://pypi.org/project/xlrd/).

## Install

```bash
go get github.com/LaoQi/xlwt-go
```

Requires a Go toolchain matching the `go` directive in [go.mod](go.mod). The
module follows standard Go versioning; import it as `xlwt`:

```go
import xlwt "github.com/LaoQi/xlwt-go"
```

## Usage

```go
package main

import (
	"log"
	"os"

	xlwt "github.com/LaoQi/xlwt-go"
)

func main() {
	wb := xlwt.NewWorkbook()
	ws, err := wb.AddSheet("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	// Row and column indices are zero-based.
	if err := ws.Write(0, 0, "Hello, XLS!"); err != nil {
		log.Fatal(err)
	}
	if err := ws.Write(1, 0, "你好，世界"); err != nil {
		log.Fatal(err)
	}
	if err := ws.Write(1, 1, "你好，世界"); err != nil { // duplicate strings share one SST entry
		log.Fatal(err)
	}

	fp, err := os.Create("output.xls")
	if err != nil {
		log.Fatal(err)
	}
	defer fp.Close()

	if err := wb.Save(fp); err != nil {
		log.Fatal(err)
	}
}
```

`Workbook.Save` accepts any `io.Writer`, so writing to a `bytes.Buffer`
(in-memory export, HTTP response body, ...) works the same way. See
[example_test.go](example_test.go) for runnable examples.

## API

| Symbol | Description |
| --- | --- |
| `NewWorkbook() *Workbook` | Create an empty workbook |
| `(*Workbook).AddSheet(name string) (*Worksheet, error)` | Append a worksheet, returns it for writing |
| `(*Worksheet).Write(row, col int, label string) error` | Write a string cell (zero-based row/col) |
| `(*Worksheet).WriteWithStyle(row, col int, label string, style *XFStyle) error` | Write a styled string cell (nil style = default) |
| `(*Workbook).AddStyle(style *XFStyle) (int, error)` | Register a style, returns its XF index |
| `(*Worksheet).SetColWidth(first, last int, characters float64) error` | Set the width of a column range |
| `(*Worksheet).SetColWidthRaw(first, last, width int) error` | Same, in 1/256 of the zero character width |
| `(*Worksheet).SetColDefaultWidth(characters int) error` | Width of columns without an explicit width |
| `(*Worksheet).SetColHidden(first, last int, hidden bool) error` | Hide or show a column range |
| `(*Worksheet).SetRowHeight(row, twips int) error` | Set a row height (1/20 point) |
| `(*Worksheet).SetRowHeightPoints(row int, points float64) error` | Same, in points |
| `(*Worksheet).SetRowDefaultHeight(twips int) error` | Height of rows without a ROW record |
| `(*Worksheet).SetRowHidden(row int, hidden bool) error` | Hide or show a row |
| `(*Workbook).Save(w io.Writer) error` | Serialize the workbook as an `.xls` stream |

### Styling

`WriteWithStyle` takes a `*XFStyle`; `nil` means the library default. A style is
a number format string plus five components:

```go
bold := xlwt.NewXFStyle()
bold.Font.Bold = true
bold.Font.Name = "Arial"
bold.Font.Height = 0x0190 // 20 points (1/20 point units)

if err := ws.WriteWithStyle(1, 0, "bold 20pt", bold); err != nil {
	log.Fatal(err)
}

money := xlwt.NewXFStyle()
money.NumFormatStr = "#,##0.00"
money.Alignment.Horz = xlwt.HorzRight

boxed := xlwt.NewXFStyle()
boxed.Borders.Bottom = xlwt.BorderThick
boxed.Pattern.Pattern = xlwt.PatternSolid
boxed.Pattern.ForeColour = 0x2C
```

| Component | Type | What it controls |
| --- | --- | --- |
| `Font` | `Font` | Typeface, size, bold, italic, underline, strike-out, colour |
| `Alignment` | `Alignment` | Horizontal/vertical alignment, wrap, shrink, indent, rotation |
| `Borders` | `Borders` | Line style and colour per edge, plus diagonals |
| `Pattern` | `Pattern` | Fill pattern, foreground and background colour |
| `Protection` | `Protection` | Cell locking and formula hiding |
| `NumFormatStr` | `string` | Number format, for example `"General"`, `"0.00%"` |

**Zero values fall back to defaults**, so a partially filled style works:

```go
// default look, bold text
style := &xlwt.XFStyle{Font: xlwt.Font{Bold: true}}
```

Two fields are exceptions, because the file format reserves 0 for a meaningful
non-default setting: `Alignment.Vert` (0 = top) and `Protection.CellUnlocked`
(0 = locked, which is also the default). Start from `NewXFStyle()` (or the
per-component constructors `NewFont`, `NewAlignment`, `NewBorders`,
`NewPattern`, `NewProtection`) to get the library defaults.

Styles are **deduplicated by value** across the workbook: reusing one style for
a million cells stores it once, and two separate but equal style objects resolve
to the same XF index. A `.xls` file can hold 4094 distinct cell styles; exceeding
that returns `ErrTooManyStyles`. Unknown number format strings are registered as
user-defined formats starting at index 164, like the Python original.

### Column widths and row heights

```go
if err := ws.SetColWidth(0, 0, 20); err != nil { // 20 characters
	log.Fatal(err)
}
if err := ws.SetColHidden(5, 5, true); err != nil {
	log.Fatal(err)
}
if err := ws.SetRowHeightPoints(0, 25); err != nil { // 25 points
	log.Fatal(err)
}
if err := ws.SetRowHidden(10, true); err != nil {
	log.Fatal(err)
}
```

Widths are measured in characters of the default font, the unit spreadsheet
applications display (`8.43` is the default). `SetColWidthRaw` takes the raw file
unit, 1/256 of the zero character width. Heights are measured in points;
`SetRowHeight` takes twips (1/20 point), the same unit as `Font.Height`.

Sizes are written only for the columns and rows you touch: a sheet that never
calls these methods produces byte-for-byte the same file as before the feature
existed. `SetColDefaultWidth` and `SetRowDefaultHeight` change the sheet-wide
defaults for columns and rows without an explicit setting.

### Validation

`Write` and `AddSheet` reject values that cannot be represented in a `.xls`
file instead of silently producing a corrupt workbook. On error nothing is
stored. Use `errors.Is` with the exported sentinels:

| Error | Cause | Limit |
| --- | --- | --- |
| `ErrRowOutOfRange` | row index outside `0..MaxRow` | `MaxRow` = 65535 |
| `ErrColOutOfRange` | column index outside `0..MaxCol` | `MaxCol` = 255 |
| `ErrStringTooLong` | more than `MaxStringLength` characters | `MaxStringLength` = 32767 |
| `ErrInvalidSheetName` | empty name, longer than `MaxSheetNameLength` characters, or containing `\ / ? * [ ] :` or a leading apostrophe | `MaxSheetNameLength` = 31 |
| `ErrDuplicateSheetName` | another sheet already uses the name (case-insensitive) | — |
| `ErrTooManyStyles` | more than 4094 distinct cell styles | — |
| `ErrTooManyNumberFormats` | too many user-defined number formats | — |
| `ErrUnattachedWorksheet` | an explicit style was passed to a worksheet built with `NewWorksheet` instead of `Workbook.AddSheet` | — |
| `ErrInvalidColumnWidth` | column width outside `0..255` characters (or `0..65535` raw) | — |
| `ErrInvalidRowHeight` | row height outside `0..32767` twips, or negative points | — |

Strings are limited in *characters*, not bytes, so multi-byte text (CJK, emoji)
is measured correctly.

## Limitations

- Only string cells; no numbers, dates, booleans or formulas
- Styling covers fonts, number formats, alignment, borders, fills and
  protection, but not the `easyxf("font:bold on")` string syntax of the Python
  original; styles are built by setting struct fields
- No cell highlighting or merged cells yet
- Outline levels for rows and columns exist in the records but are not exposed
  through the API
- Write-only: the library cannot read `.xls` files
- Cells are collected in memory until `Save` is called; there is no streaming
  or incremental writer

## Implementation notes

The package builds the BIFF8 record stream (`biff.go`, `record_*.go`,
`workbook.go`, `worksheet.go`, `sst.go`, `style.go`) and wraps it into an OLE2
compound document (`compound_doc.go`) with 512-byte sectors. Type aliases in
`types.go` mirror the `struct.pack` layout of the Python original, so byte
layouts are intentionally kept 1:1 with upstream `xlwt`.

## Testing

```bash
go build ./...
go vet ./...
go test -race ./...
```

Tests write into a temporary directory. Besides structural assertions (OLE2
signature, sector alignment) they include a small BIFF parser that validates
content level details: shared string table layout and CONTINUE splitting,
BOUNDSHEET offsets pointing at their own worksheet, and the worksheet that each
cell belongs to. The generated files have additionally been verified with
Python `xlrd` 2.0.1 (multi-sheet workbooks with CJK sheet names, 5000-character
and emoji strings, and a 33 MB workbook exercising the secondary MSAT sectors).

## License

[LGPL-2.1](LICENSE), same as the upstream Python xlwt library.