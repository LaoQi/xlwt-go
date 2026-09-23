package xlwt

import (
	"bytes"
	"fmt"
	"sort"
)

type Cell struct {
	Row    int
	Col    int
	SSTIdx int
	XFIdx  int
}

type Worksheet struct {
	Name      string
	SST       *SharedStringTable
	Grid      map[uint32]Cell
	RowsIndex map[int]bool

	// Cols and Rows hold the explicit column and row sizing. Only the entries
	// that were set explicitly are recorded and written to the file.
	Cols map[int]colInfo
	Rows map[int]rowInfo

	// ColDefaultWidth and RowDefaultHeight hold the sheet-wide defaults set via
	// SetColDefaultWidth and SetRowDefaultHeight. Zero means "file format
	// default" and produces no extra record.
	ColDefaultWidth  int
	RowDefaultHeight int

	// workbook is the workbook this sheet was added to; it owns the shared
	// string table and the style collection.
	workbook *Workbook
}

// NewWorksheet creates a worksheet. It is used internally by
// [Workbook.AddSheet], which also links the sheet to its workbook; a worksheet
// created directly cannot be written to because it has no style collection to
// register styles with.
func NewWorksheet(name string, sst *SharedStringTable) *Worksheet {
	return &Worksheet{
		Name:      name,
		SST:       sst,
		Grid:      make(map[uint32]Cell),
		RowsIndex: make(map[int]bool),
		Cols:      make(map[int]colInfo),
		Rows:      make(map[int]rowInfo),
	}
}

func (ws *Worksheet) calcSettingsRec() []byte {
	var buf bytes.Buffer
	buf.Write(CalcModeRecord(1))
	buf.Write(CalcCountRecord(0x0064))
	buf.Write(RefModeRecord(1))
	buf.Write(IterationRecord(0))
	buf.Write(DeltaRecord(0.001))
	buf.Write(SaveRecalcRecord(0))

	return buf.Bytes()
}

// GutsRec writes the GUTS record. The outline level counts follow the Python
// original, which reports one visible row level even when no outlines are used.
func (ws *Worksheet) GutsRec() []byte {
	rowLevels := ws.rowVisibleLevels()
	if rowLevels == 0 {
		rowLevels = 1
	}
	return GutsRecord(0, 0, rowLevels, ws.colVisibleLevels())
}

func (ws *Worksheet) wsBoolRec() []byte {
	options := 0x01
	options |= 0x01 << 10 // __show_row_outline
	options |= 0x01 << 11 // __show_col_outline
	return WSBoolRecord(options)
}

// remapSSTIndexes rewrites every cell of the sheet so that Cell.SSTIdx keeps
// pointing at its own string after the shared string table was sorted. It is
// called by Workbook.finalizeSST and is the only place that mutates a stored
// cell other than Write.
func (ws *Worksheet) remapSSTIndexes(remap []int) {
	for key, cell := range ws.Grid {
		if cell.SSTIdx >= 0 && cell.SSTIdx < len(remap) {
			cell.SSTIdx = remap[cell.SSTIdx]
			ws.Grid[key] = cell
		}
	}
}

func (ws *Worksheet) dimensionsRec() []byte {
	lastUsedRow := 0
	lastUsedCol := 0
	for _, cell := range ws.Grid {
		if cell.Row > lastUsedRow {
			lastUsedRow = cell.Row
		}
		if cell.Col > lastUsedCol {
			lastUsedCol = cell.Col
		}
	}
	return DimensionsRecord(0, lastUsedRow+1, 0, lastUsedCol+1)
}

func (ws *Worksheet) printSettingsRec() []byte {
	var buf bytes.Buffer
	buf.Write(PrintHeadersRecord(0))
	buf.Write(PrintGridLinesRecord(0))
	buf.Write(GridSetRecord(1))
	buf.Write(HorizontalPageBreaksRecord())
	buf.Write(VerticalPageBreaksRecord())
	buf.Write(HeaderRecord("&P"))
	buf.Write(FooterRecord("&F"))
	buf.Write(HCenterRecord(1))
	buf.Write(VCenterRecord(0))
	buf.Write(LeftMarginRecord(0.3))
	buf.Write(RightMarginRecord(0.3))
	buf.Write(TopMarginRecord(0.61))
	buf.Write(BottomMarginRecord(0.37))
	buf.Write(SetupPageRecord())

	return buf.Bytes()
}

func (ws *Worksheet) protectionRec() []byte {
	var buf bytes.Buffer
	buf.Write(ProtectRecord(0))
	buf.Write(ScenProtectRecord(0))
	buf.Write(WindowProtectRecord(0))
	buf.Write(ObjectProtectRecord(0))
	buf.Write(PasswordRecord(""))
	return buf.Bytes()
}

func (ws *Worksheet) GetRowCellsBiffData(row int) []byte {
	// The whole grid has to be scanned for a single row, so this entry point is
	// only cheap when the caller already holds the cells of that row (it is kept
	// for compatibility).
	return ws.rowCellsBiffData(row, ws.cellsOfRow(row))
}

// rowCellsBiffData writes one row from cells that are known to belong to it.
func (ws *Worksheet) rowCellsBiffData(row int, cells []Cell) []byte {
	if len(cells) == 0 {
		// A row without cells still needs a ROW record when its properties were
		// set explicitly (height, hidden, outline level).
		if _, explicit := ws.Rows[row]; !explicit {
			return []byte{}
		}
		var buf bytes.Buffer
		buf.Write(ws.rowRec(row, 0, 0))
		return buf.Bytes()
	}
	sort.Slice(cells, func(i, j int) bool {
		return cells[i].Col < cells[j].Col
	})
	firstCol := cells[0].Col
	lastCol := cells[len(cells)-1].Col + 1

	var buf bytes.Buffer
	buf.Write(ws.rowRec(row, firstCol, lastCol))

	for _, cell := range cells {
		buf.Write(LabelSSTRecord(row, cell.Col, cell.XFIdx, cell.SSTIdx))
	}

	return buf.Bytes()
}

func (ws *Worksheet) GetRowsBiffData() []byte {
	// Rows are emitted when they contain cells or when their sizing was set
	// explicitly, which is what the Python original does for row().hidden etc.
	byRow := ws.cellsByRow()

	seen := make(map[int]bool, len(byRow)+len(ws.Rows))
	for index := range byRow {
		seen[index] = true
	}
	for index := range ws.Rows {
		seen[index] = true
	}

	rows := make([]int, 0, len(seen))
	for index := range seen {
		rows = append(rows, index)
	}
	sort.Ints(rows)

	var buf bytes.Buffer
	for _, index := range rows {
		buf.Write(ws.rowCellsBiffData(index, byRow[index]))
	}
	return buf.Bytes()
}

// cellsOfRow returns the cells of one row. It is the single-row convenience
// wrapper around cellsByRow and is as expensive as a full grid scan.
func (ws *Worksheet) cellsOfRow(row int) []Cell {
	var cells []Cell
	for _, cell := range ws.Grid {
		if cell.Row == row {
			cells = append(cells, cell)
		}
	}
	return cells
}

// cellsByRow groups the grid by row. The rows are written row by row, so
// collecting the cells once keeps serialisation linear in the number of cells;
// scanning the whole grid for every row would make it quadratic.
func (ws *Worksheet) cellsByRow() map[int][]Cell {
	byRow := make(map[int][]Cell, len(ws.RowsIndex))
	for _, cell := range ws.Grid {
		byRow[cell.Row] = append(byRow[cell.Row], cell)
	}
	return byRow
}

func (ws *Worksheet) GetBiffData() []byte {
	var buf bytes.Buffer
	buf.Write(Biff8BOFRecord(Biff8BOFRecord__WORKSHEET))
	buf.Write(ws.calcSettingsRec())
	buf.Write(ws.GutsRec())
	buf.Write(ws.defaultRowHeightRec())
	buf.Write(ws.wsBoolRec())
	buf.Write(ws.colInfoRec())
	buf.Write(ws.defColWidthRec())
	buf.Write(ws.dimensionsRec())
	buf.Write(ws.printSettingsRec())
	buf.Write(ws.protectionRec())

	buf.Write(ws.GetRowsBiffData())

	// skip MergedCellsRecord ObjBmpRecord
	buf.Write(DefaultWindow2Record())
	// skip PanesRecord
	buf.Write(EOFRecord())

	return buf.Bytes()
}

// Write stores a string cell at the zero-based row r and column c.
//
// It returns [ErrRowOutOfRange], [ErrColOutOfRange] or [ErrStringTooLong] when
// the arguments cannot be represented in a BIFF8 file, in which case the
// worksheet is left unchanged. Writing the same cell twice keeps the last value.
//
// It is shorthand for WriteWithStyle without a style.
func (ws *Worksheet) Write(r, c int, label string) error {
	return ws.WriteWithStyle(r, c, label, nil)
}

// WriteWithStyle stores a string cell with an explicit style. A nil style means
// the library default style.
//
// Styles are deduplicated by value across the whole workbook, so a style can be
// reused for any number of cells. Passing more distinct styles than a BIFF8 file
// can reference fails with [ErrTooManyStyles].
func (ws *Worksheet) WriteWithStyle(r, c int, label string, style *XFStyle) error {
	if r < 0 || r > MaxRow {
		return fmt.Errorf("%w: row %d, allowed range is 0..%d", ErrRowOutOfRange, r, MaxRow)
	}
	if c < 0 || c > MaxCol {
		return fmt.Errorf("%w: column %d, allowed range is 0..%d", ErrColOutOfRange, c, MaxCol)
	}
	if len([]rune(label)) > MaxStringLength {
		return fmt.Errorf("%w: %d characters, allowed maximum is %d", ErrStringTooLong, len([]rune(label)), MaxStringLength)
	}

	// A worksheet created through Workbook.AddSheet registers the style with its
	// workbook. A worksheet built directly with NewWorksheet has no style
	// collection, so it can only use the default style.
	xfIdx := DefaultCellXFStyle
	if ws.workbook != nil {
		idx, err := ws.workbook.AddStyle(style)
		if err != nil {
			return err
		}
		xfIdx = idx
	} else if style != nil {
		return fmt.Errorf("%w: worksheet %q was not created by Workbook.AddSheet", ErrUnattachedWorksheet, ws.Name)
	}

	key := uint32((r << 16) + c)
	idx := ws.SST.AddStr(label)
	ws.Grid[key] = Cell{Row: r, Col: c, SSTIdx: idx, XFIdx: xfIdx}
	ws.RowsIndex[r] = true
	return nil
}
