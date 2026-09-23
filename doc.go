// Package xlwt writes Microsoft Excel 97-2003 spreadsheets (.xls).
//
// The package emits the BIFF8 record stream and wraps it in an OLE2 compound
// document, producing files readable by Excel, WPS, LibreOffice and tools such
// as Python's xlrd. It is a pure Go port of the Python xlwt library
// (https://github.com/python-excel/xlwt) and depends only on the standard
// library.
//
// Basic usage:
//
//	wb := xlwt.NewWorkbook()
//	ws, err := wb.AddSheet("Sheet1")
//	if err != nil {
//		return err
//	}
//	if err := ws.Write(0, 0, "Hello, XLS!"); err != nil {
//		return err
//	}
//	if err := wb.Save(w); err != nil {
//		return err
//	}
//
// Row and column indices passed to [Worksheet.Write] are zero-based. Write and
// AddSheet validate their arguments against the limits of the file format and
// return errors wrapping [ErrRowOutOfRange], [ErrColOutOfRange],
// [ErrStringTooLong], [ErrInvalidSheetName] or [ErrDuplicateSheetName]; on error
// nothing is stored, so a workbook never contains data that .xls cannot
// represent.
//
// Cells can be styled individually:
//
//	bold := xlwt.NewXFStyle()
//	bold.Font.Bold = true
//	if err := ws.WriteWithStyle(1, 0, "bold", bold); err != nil {
//		return err
//	}
//
// Styles are deduplicated by value across the workbook, so one style object can
// be reused for any number of cells. A partially filled style falls back to the
// documented defaults of the individual components; see [Font], [Alignment],
// [Borders], [Pattern] and [Protection].
//
// Column widths and row heights can be set explicitly; a sheet that does not
// use them is written exactly as before:
//
//	if err := ws.SetColWidth(0, 0, 20); err != nil { // 20 characters
//		return err
//	}
//	if err := ws.SetRowHeightPoints(0, 25); err != nil { // 25 points
//		return err
//	}
//
// Widths are given in characters of the default font and heights in points,
// matching what spreadsheets display. See SetColWidthRaw and SetRowHeight for
// the raw file units.
//
// The bytes of a workbook do not depend on the order in which the cells were
// written. Shared string indexes follow the first cell that refers to each
// string, in sheet, row and then column order, so writing the cells in that
// order (or sorting them before writing) keeps the historical byte-for-byte
// layout, while any other order produces exactly the same file as well. A
// caller that fills a sheet from a map therefore no longer gets a different
// file on every run.
//
// This is a write-only library; it cannot read .xls files. Only string cells are
// supported. See the project README for the current limitations.
package xlwt
