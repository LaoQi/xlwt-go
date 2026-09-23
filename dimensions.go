package xlwt

import (
	"fmt"
	"math"
	"sort"
)

// Column and row sizing, mirroring the Python original's Column and Row classes.
//
// Units follow the file format:
//
//   - column width is stored in 1/256 of the width of the "0" character of the
//     default font. The Excel user interface shows it in characters, so the
//     helpers convert: widthIn256ths = int(characters * 256).
//   - row height is stored in twips (1/20 of a point), the same unit as
//     [Font].Height.
//
// Neither is written to the file unless it is set explicitly; a sheet without
// explicit sizes produces exactly the same bytes as before this feature existed.
const (
	// DefaultColumnWidth is the width a column gets when a COLINFO record is
	// written without an explicit width: 0x0B92 / 256 = 11.57 characters, the
	// value the Python original uses for its Column objects.
	DefaultColumnWidth = 0x0B92

	// DefaultColumnWidthChars is the user interface default (8.43 characters)
	// that the file format assumes for columns without a COLINFO record.
	DefaultColumnWidthChars = 0x0008

	// DefaultRowHeight is the default height of rows without a ROW record, in
	// twips: 0x00FF = 12.75 points.
	DefaultRowHeight = 0x00FF

	// rowDefaultXFIndex is the XF index written into the ROW and COLINFO records
	// for cells that have no cell record of their own. It matches the Python
	// original, which uses 0x0F for both.
	rowDefaultXFIndex = 0x0F
)

// colInfo describes a column range as stored in a COLINFO record.
type colInfo struct {
	width  int
	hidden bool
	// level is the outline level (0..7) and collapse the collapsed flag.
	level    int
	collapse bool
	// bestFit asks the application to auto-fit the column.
	bestFit bool
	// userSet marks the width as set by the user.
	userSet bool
}

// options packs the COLINFO option flags.
func (c colInfo) options() int {
	options := 0
	if c.hidden {
		options |= 0x01 << 0
	}
	if c.userSet {
		options |= 0x01 << 1
	}
	if c.bestFit {
		options |= 0x01 << 2
	}
	options |= (c.level & 0x07) << 8
	if c.collapse {
		options |= 0x01 << 12
	}
	return options
}

// rowInfo describes the sizing of a single row.
type rowInfo struct {
	height   int
	hidden   bool
	level    int
	collapse bool
	// spaceAbove/spaceBelow add padding for thick cell borders.
	spaceAbove bool
	spaceBelow bool
}

// heightOptions packs the height word of a ROW record.
//
// Bit 15 ("row has default height") is left clear, matching the Python original,
// which keeps has_default_height at zero for every row including those that use
// the default height.
func (r rowInfo) heightOptions() int {
	return r.height & 0x07FFF
}

// options packs the option word of a ROW record.
func (r rowInfo) options() int {
	options := (r.level & 0x07) << 0
	if r.collapse {
		options |= 0x01 << 4
	}
	if r.hidden {
		options |= 0x01 << 5
	}
	// Bit 8 is always set. Bits 27-16 carry the XF index used by cells that have
	// no cell record; the Python original writes 0x0F there.
	options |= 0x01 << 8
	options |= (rowDefaultXFIndex & 0x0FFF) << 16
	if r.spaceAbove {
		options |= 0x01 << 28
	}
	if r.spaceBelow {
		options |= 0x01 << 29
	}
	return options
}

// SetColWidth sets the width of the column range firstCol..lastCol, measured in
// characters of the default font (the unit the Excel user interface shows, for
// example 8.43).
func (ws *Worksheet) SetColWidth(firstCol, lastCol int, characters float64) error {
	if err := ws.checkColRange(firstCol, lastCol); err != nil {
		return err
	}
	if characters < 0 || characters > 255 {
		return fmt.Errorf("%w: column width %v characters, allowed range is 0..255", ErrInvalidColumnWidth, characters)
	}
	width := int(math.Round(characters * 256))
	for c := firstCol; c <= lastCol; c++ {
		info := ws.colInfoFor(c)
		info.width = width
		info.userSet = true
		ws.Cols[c] = info
	}
	return nil
}

// SetColWidthRaw sets the width of the column range using the raw file format
// unit: 1/256 of the width of the "0" character.
func (ws *Worksheet) SetColWidthRaw(firstCol, lastCol, width int) error {
	if err := ws.checkColRange(firstCol, lastCol); err != nil {
		return err
	}
	if width < 0 || width > 65535 {
		return fmt.Errorf("%w: raw width %d, allowed range is 0..65535", ErrInvalidColumnWidth, width)
	}
	for c := firstCol; c <= lastCol; c++ {
		info := ws.colInfoFor(c)
		info.width = width
		info.userSet = true
		ws.Cols[c] = info
	}
	return nil
}

// SetColHidden hides or shows the column range firstCol..lastCol.
func (ws *Worksheet) SetColHidden(firstCol, lastCol int, hidden bool) error {
	if err := ws.checkColRange(firstCol, lastCol); err != nil {
		return err
	}
	for c := firstCol; c <= lastCol; c++ {
		info := ws.colInfoFor(c)
		info.hidden = hidden
		ws.Cols[c] = info
	}
	return nil
}

// SetColDefaultWidth sets the width used for columns without an explicit width,
// by writing a DEFCOLWIDTH record. Zero restores the file format default.
func (ws *Worksheet) SetColDefaultWidth(characters int) error {
	if characters < 0 || characters > 65535 {
		return fmt.Errorf("%w: default width %d, allowed range is 0..65535", ErrInvalidColumnWidth, characters)
	}
	ws.ColDefaultWidth = characters
	return nil
}

// SetRowHeight sets the height of a row in twips (1/20 of a point), the same
// unit as Font.Height. Use SetRowHeightPoints for a friendlier unit.
func (ws *Worksheet) SetRowHeight(row, twips int) error {
	if err := ws.checkRow(row); err != nil {
		return err
	}
	if twips < 0 || twips > 0x7FFF {
		return fmt.Errorf("%w: row height %d twips, allowed range is 0..%d", ErrInvalidRowHeight, twips, 0x7FFF)
	}
	info := ws.rowInfoFor(row)
	info.height = twips
	ws.Rows[row] = info
	return nil
}

// SetRowHeightPoints sets the height of a row in points (1 point = 20 twips),
// which is the unit spreadsheets show.
func (ws *Worksheet) SetRowHeightPoints(row int, points float64) error {
	if points < 0 {
		return fmt.Errorf("%w: row height %v points must not be negative", ErrInvalidRowHeight, points)
	}
	return ws.SetRowHeight(row, int(math.Round(points*20)))
}

// SetRowHidden hides or shows a row.
func (ws *Worksheet) SetRowHidden(row int, hidden bool) error {
	if err := ws.checkRow(row); err != nil {
		return err
	}
	info := ws.rowInfoFor(row)
	info.hidden = hidden
	ws.Rows[row] = info
	return nil
}

// SetRowDefaultHeight sets the height of rows without a ROW record, in twips.
func (ws *Worksheet) SetRowDefaultHeight(twips int) error {
	if twips < 0 || twips > 0x7FFF {
		return fmt.Errorf("%w: default row height %d twips, allowed range is 0..%d", ErrInvalidRowHeight, twips, 0x7FFF)
	}
	ws.RowDefaultHeight = twips
	return nil
}

// colInfoFor returns the current description of a column, or its default.
func (ws *Worksheet) colInfoFor(col int) colInfo {
	if info, ok := ws.Cols[col]; ok {
		return info
	}
	return colInfo{width: DefaultColumnWidth}
}

// rowInfoFor returns the current description of a row, or its default.
func (ws *Worksheet) rowInfoFor(row int) rowInfo {
	if info, ok := ws.Rows[row]; ok {
		return info
	}
	return rowInfo{height: DefaultRowHeight}
}

func (ws *Worksheet) checkColRange(firstCol, lastCol int) error {
	if firstCol < 0 || firstCol > MaxCol {
		return fmt.Errorf("%w: column %d, allowed range is 0..%d", ErrColOutOfRange, firstCol, MaxCol)
	}
	if lastCol < firstCol || lastCol > MaxCol {
		return fmt.Errorf("%w: last column %d, want %d..%d", ErrColOutOfRange, lastCol, firstCol, MaxCol)
	}
	return nil
}

func (ws *Worksheet) checkRow(row int) error {
	if row < 0 || row > MaxRow {
		return fmt.Errorf("%w: row %d, allowed range is 0..%d", ErrRowOutOfRange, row, MaxRow)
	}
	return nil
}

// colInfoRec serialises one COLINFO record per column that has an explicit
// width or visibility setting, in column order.
func (ws *Worksheet) colInfoRec() []byte {
	if len(ws.Cols) == 0 {
		return nil
	}

	columns := make([]int, 0, len(ws.Cols))
	for col := range ws.Cols {
		columns = append(columns, col)
	}
	sort.Ints(columns)

	var buf []byte
	for _, col := range columns {
		info := ws.Cols[col]
		buf = append(buf, ColInfoRecord(col, col, info.width, rowDefaultXFIndex, info.options(), 0)...)
	}
	return buf
}

// defaultRowHeightRec writes the DEFAULTROWHEIGHT record. When the default height
// is unchanged this is the same record the library always produced.
func (ws *Worksheet) defaultRowHeightRec() []byte {
	height := ws.RowDefaultHeight
	if height == 0 {
		height = DefaultRowHeight
	}
	return DefaultRowHeightRecord(0x0000, height)
}

// defColWidthRec writes the DEFCOLWIDTH record, or nothing when the default
// width was not changed.
func (ws *Worksheet) defColWidthRec() []byte {
	if ws.ColDefaultWidth == 0 {
		return nil
	}
	return DefColWidthRecord(ws.ColDefaultWidth)
}

// rowRec serialises the ROW record of a row.
func (ws *Worksheet) rowRec(row, firstCol, lastCol int) []byte {
	info := ws.rowInfoFor(row)
	return RowRecord(row, firstCol, lastCol, info.heightOptions(), info.options())
}

// colVisibleLevels returns the number of column outline levels in use, which the
// GUTS record reports.
func (ws *Worksheet) colVisibleLevels() int {
	levels := 0
	for _, info := range ws.Cols {
		if info.level+1 > levels {
			levels = info.level + 1
		}
	}
	return levels
}

// rowVisibleLevels returns the number of row outline levels in use.
func (ws *Worksheet) rowVisibleLevels() int {
	levels := 0
	for _, info := range ws.Rows {
		if info.level+1 > levels {
			levels = info.level + 1
		}
	}
	return levels
}
