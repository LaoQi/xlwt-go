package xlwt

import (
	"encoding/binary"
	"testing"
)

const colInfoID = 0x007D
const defColWidthID = 0x0055
const defaultRowHeightID = 0x0225
const rowID = 0x0208

// colInfoRecords extracts every COLINFO record payload, in order.
func colInfoRecords(t *testing.T, stream []byte) [][]byte {
	t.Helper()
	var out [][]byte
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) == colInfoID {
			out = append(out, rec[1].([]byte))
		}
	}
	return out
}

// rowRecords extracts every ROW record payload, keyed by row index.
func rowRecords(t *testing.T, stream []byte) map[int][]byte {
	t.Helper()
	out := map[int][]byte{}
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) == rowID {
			data := rec[1].([]byte)
			out[int(binary.LittleEndian.Uint16(data[0:2]))] = data
		}
	}
	return out
}

// defaultRowHeight returns the height word of the DEFAULTROWHEIGHT record.
func defaultRowHeight(t *testing.T, stream []byte) (options, height int) {
	t.Helper()
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) == defaultRowHeightID {
			data := rec[1].([]byte)
			return int(binary.LittleEndian.Uint16(data[0:2])), int(binary.LittleEndian.Uint16(data[2:4]))
		}
	}
	t.Fatal("no DEFAULTROWHEIGHT record")
	return 0, 0
}

// TestDimensions_UnaffectedByDefault is the guardrail: a sheet that never sets a
// size must not contain COLINFO or DEFCOLWIDTH records and must keep the default
// row height, so its bytes stay identical to a workbook written before column
// and row sizing existed.
func TestDimensions_UnaffectedByDefault(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(0, 0, "plain"); err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(1, 1, "中文"); err != nil {
		t.Fatal(err)
	}

	stream := wb.GetBiffData()

	if recs := colInfoRecords(t, stream); len(recs) != 0 {
		t.Errorf("COLINFO records = %d, want 0 when no width was set", len(recs))
	}
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) == defColWidthID {
			t.Error("DEFCOLWIDTH record written although no default width was set")
		}
	}

	options, height := defaultRowHeight(t, stream)
	if options != 0x0000 {
		t.Errorf("DEFAULTROWHEIGHT options = 0x%04X, want 0x0000", options)
	}
	if height != DefaultRowHeight {
		t.Errorf("DEFAULTROWHEIGHT height = %d, want %d", height, DefaultRowHeight)
	}

	// The ROW records must keep the historical layout: default height, bit 15
	// clear, XF index 0x0F in bits 27-16 and bit 8 set.
	rows := rowRecords(t, stream)
	if len(rows) != 2 {
		t.Fatalf("ROW records = %d, want 2", len(rows))
	}
	for idx, data := range rows {
		heightWord := int(binary.LittleEndian.Uint16(data[6:8]))
		if heightWord != DefaultRowHeight {
			t.Errorf("row %d height word = 0x%04X, want 0x%04X", idx, heightWord, DefaultRowHeight)
		}
		if heightWord&0x8000 != 0 {
			t.Errorf("row %d has the 'default height' bit set; the Python original leaves it clear", idx)
		}
		optionsWord := binary.LittleEndian.Uint32(data[12:16])
		if got := (optionsWord >> 16) & 0x0FFF; got != rowDefaultXFIndex {
			t.Errorf("row %d default XF index = 0x%X, want 0x%X", idx, got, rowDefaultXFIndex)
		}
		if optionsWord&(1<<8) == 0 {
			t.Errorf("row %d does not have the always-set bit 8", idx)
		}
	}
}

// TestDimensions_ColWidth checks the character to 1/256 conversion and the
// COLINFO option flags.
func TestDimensions_ColWidth(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(0, 0, "x"); err != nil {
		t.Fatal(err)
	}

	if err := ws.SetColWidth(3, 3, 20); err != nil {
		t.Fatal(err)
	}
	if err := ws.SetColWidth(5, 6, 8.43); err != nil {
		t.Fatal(err)
	}
	if err := ws.SetColHidden(4, 4, true); err != nil {
		t.Fatal(err)
	}

	recs := colInfoRecords(t, wb.GetBiffData())
	if len(recs) != 4 {
		t.Fatalf("COLINFO records = %d, want 4 (columns 3, 4, 5, 6)", len(recs))
	}

	// COLINFO records are written in column order.
	type colRec struct{ first, last, width, xf, options int }
	var parsed []colRec
	for _, data := range recs {
		parsed = append(parsed, colRec{
			first:   int(binary.LittleEndian.Uint16(data[0:2])),
			last:    int(binary.LittleEndian.Uint16(data[2:4])),
			width:   int(binary.LittleEndian.Uint16(data[4:6])),
			xf:      int(binary.LittleEndian.Uint16(data[6:8])),
			options: int(binary.LittleEndian.Uint16(data[8:10])),
		})
	}

	if parsed[0].first != 3 || parsed[0].last != 3 {
		t.Errorf("first COLINFO range = %d..%d, want 3..3", parsed[0].first, parsed[0].last)
	}
	if want := 20 * 256; parsed[0].width != want {
		t.Errorf("column 3 width = %d, want %d (20 characters)", parsed[0].width, want)
	}
	if parsed[0].options&0x02 == 0 {
		t.Error("column 3 does not have the user-set flag")
	}
	if parsed[0].xf != rowDefaultXFIndex {
		t.Errorf("column 3 XF index = 0x%X, want 0x%X", parsed[0].xf, rowDefaultXFIndex)
	}

	if parsed[1].first != 4 || parsed[1].options&0x01 == 0 {
		t.Errorf("column 4 must be hidden: range %d..%d options 0x%X", parsed[1].first, parsed[1].last, parsed[1].options)
	}

	// 8.43 characters rounds to 2158 in 1/256 units.
	if want := 2158; parsed[2].width != want {
		t.Errorf("column 5 width = %d, want %d (8.43 characters)", parsed[2].width, want)
	}
	if parsed[3].first != 6 || parsed[3].width != parsed[2].width {
		t.Errorf("column 6 range/width = %d..%d/%d, want 6..6/%d", parsed[3].first, parsed[3].last, parsed[3].width, parsed[2].width)
	}
}

// TestDimensions_ColWidthRaw checks the raw unit setter.
func TestDimensions_ColWidthRaw(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	if err := ws.SetColWidthRaw(0, 2, 2962); err != nil {
		t.Fatal(err)
	}

	recs := colInfoRecords(t, wb.GetBiffData())
	if len(recs) != 3 {
		t.Fatalf("COLINFO records = %d, want 3", len(recs))
	}
	for i, data := range recs {
		if got := int(binary.LittleEndian.Uint16(data[4:6])); got != 2962 {
			t.Errorf("column %d width = %d, want 2962", i, got)
		}
		if got := int(binary.LittleEndian.Uint16(data[0:2])); got != i {
			t.Errorf("COLINFO %d covers column %d, want %d", i, got, i)
		}
	}
}

// TestDimensions_RowHeight checks row heights in twips and points and the hidden
// flag.
func TestDimensions_RowHeight(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(0, 0, "a"); err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(2, 0, "b"); err != nil {
		t.Fatal(err)
	}

	if err := ws.SetRowHeight(0, 500); err != nil { // 25 points in twips
		t.Fatal(err)
	}
	if err := ws.SetRowHeightPoints(2, 25); err != nil {
		t.Fatal(err)
	}

	rows := rowRecords(t, wb.GetBiffData())
	for _, idx := range []int{0, 2} {
		data, ok := rows[idx]
		if !ok {
			t.Fatalf("no ROW record for row %d", idx)
		}
		if got := int(binary.LittleEndian.Uint16(data[6:8])); got != 500 {
			t.Errorf("row %d height = %d, want 500 twips", idx, got)
		}
	}

	// Hiding a row is independent of its height, and a hidden row without cells
	// still gets a ROW record.
	if err := ws.SetRowHidden(7, true); err != nil {
		t.Fatal(err)
	}
	data, ok := rowRecords(t, wb.GetBiffData())[7]
	if !ok {
		t.Fatal("no ROW record for the hidden row 7")
	}
	options := binary.LittleEndian.Uint32(data[12:16])
	if options&(1<<5) == 0 {
		t.Error("row 7 does not have the hidden flag (bit 5)")
	}
}

// TestDimensions_RowDefaultHeight checks the DEFAULTROWHEIGHT record.
func TestDimensions_RowDefaultHeight(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}

	if err := ws.SetRowDefaultHeight(320); err != nil { // 16 points
		t.Fatal(err)
	}
	options, height := defaultRowHeight(t, wb.GetBiffData())
	if options != 0x0000 {
		t.Errorf("options = 0x%04X, want 0x0000", options)
	}
	if height != 320 {
		t.Errorf("height = %d, want 320", height)
	}
}

// TestDimensions_ColDefaultWidth checks the DEFCOLWIDTH record is only written
// when a default width was requested.
func TestDimensions_ColDefaultWidth(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(0, 0, "x"); err != nil {
		t.Fatal(err)
	}

	if err := ws.SetColDefaultWidth(10); err != nil {
		t.Fatal(err)
	}

	var found bool
	for _, rec := range parseBiffRecords(t, wb.GetBiffData()) {
		if rec[0].(uint16) == defColWidthID {
			found = true
			if got := int(binary.LittleEndian.Uint16(rec[1].([]byte)[0:2])); got != 10 {
				t.Errorf("DEFCOLWIDTH = %d, want 10", got)
			}
		}
	}
	if !found {
		t.Error("no DEFCOLWIDTH record although a default width was set")
	}
}

// TestDimensions_GutsLevels checks that outline levels reported by the GUTS
// record follow the Python original: at least one row level, and column levels
// derived from the COLINFO records in use.
func TestDimensions_GutsLevels(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(0, 0, "x"); err != nil {
		t.Fatal(err)
	}

	// Without outline levels the row level count is still 1, as upstream does.
	guts := func() (rowLevels, colLevels int) {
		for _, rec := range parseBiffRecords(t, wb.GetBiffData()) {
			if rec[0].(uint16) == 0x0080 {
				data := rec[1].([]byte)
				return int(binary.LittleEndian.Uint16(data[4:6])), int(binary.LittleEndian.Uint16(data[6:8]))
			}
		}
		t.Fatal("no GUTS record")
		return 0, 0
	}

	if r, c := guts(); r != 1 || c != 0 {
		t.Errorf("GUTS levels = %d/%d, want 1/0 with no outlines", r, c)
	}

	// A column at outline level 2 makes the column level count 3.
	ws.Cols[0] = colInfo{width: DefaultColumnWidth, level: 2}
	if r, c := guts(); r != 1 || c != 3 {
		t.Errorf("GUTS levels = %d/%d, want 1/3 with a level-2 column", r, c)
	}

	// A row at outline level 1 makes the row level count 2.
	ws.Rows[0] = rowInfo{height: DefaultRowHeight, level: 1}
	if r, c := guts(); r != 2 || c != 3 {
		t.Errorf("GUTS levels = %d/%d, want 2/3", r, c)
	}
}

// TestDimensions_Validation covers the argument checks.
func TestDimensions_Validation(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}

	cases := []struct {
		name    string
		run     func() error
		wantErr error
	}{
		{"negative column", func() error { return ws.SetColWidth(-1, 0, 10) }, ErrColOutOfRange},
		{"column past limit", func() error { return ws.SetColWidth(0, MaxCol+1, 10) }, ErrColOutOfRange},
		{"inverted range", func() error { return ws.SetColWidth(5, 4, 10) }, ErrColOutOfRange},
		{"negative width", func() error { return ws.SetColWidth(0, 0, -1) }, ErrInvalidColumnWidth},
		{"width too large", func() error { return ws.SetColWidth(0, 0, 256) }, ErrInvalidColumnWidth},
		{"raw width too large", func() error { return ws.SetColWidthRaw(0, 0, 65536) }, ErrInvalidColumnWidth},
		{"negative row", func() error { return ws.SetRowHeight(-1, 300) }, ErrRowOutOfRange},
		{"row past limit", func() error { return ws.SetRowHeight(MaxRow+1, 300) }, ErrRowOutOfRange},
		{"height too large", func() error { return ws.SetRowHeight(0, 0x8000) }, ErrInvalidRowHeight},
		{"negative points", func() error { return ws.SetRowHeightPoints(0, -1) }, ErrInvalidRowHeight},
		{"default height too large", func() error { return ws.SetRowDefaultHeight(0x8000) }, ErrInvalidRowHeight},
		{"default width too large", func() error { return ws.SetColDefaultWidth(65536) }, ErrInvalidColumnWidth},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			err := tc.run()
			if err == nil {
				t.Fatalf("got nil, want %v", tc.wantErr)
			}
			if !errorsIs(err, tc.wantErr) {
				t.Errorf("got %v, want it to wrap %v", err, tc.wantErr)
			}
		})
	}

	// A failed call must not leave state behind.
	if len(ws.Cols) != 0 || len(ws.Rows) != 0 {
		t.Errorf("failed calls modified the sheet: Cols=%v Rows=%v", ws.Cols, ws.Rows)
	}
}

// TestDimensions_MultiSheetIsolation checks that sizing stays per worksheet.
func TestDimensions_MultiSheetIsolation(t *testing.T) {
	wb := NewWorkbook()
	wide, err := wb.AddSheet("Wide")
	if err != nil {
		t.Fatal(err)
	}
	plain, err := wb.AddSheet("Plain")
	if err != nil {
		t.Fatal(err)
	}
	if err := wide.Write(0, 0, "x"); err != nil {
		t.Fatal(err)
	}
	if err := plain.Write(0, 0, "x"); err != nil {
		t.Fatal(err)
	}
	if err := wide.SetColWidth(0, 0, 30); err != nil {
		t.Fatal(err)
	}
	if err := wide.SetRowHeightPoints(0, 40); err != nil {
		t.Fatal(err)
	}

	stream := wb.GetBiffData()
	sheets := boundSheets(t, stream)
	if len(sheets) != 2 {
		t.Fatalf("BOUNDSHEET records = %d, want 2", len(sheets))
	}

	// Slice the stream into the two worksheet substreams.
	wideStream := stream[sheets[0].Pos:sheets[1].Pos]
	plainStream := stream[sheets[1].Pos:]

	if got := len(colInfoRecords(t, wideStream)); got != 1 {
		t.Errorf("wide sheet COLINFO records = %d, want 1", got)
	}
	if got := len(colInfoRecords(t, plainStream)); got != 0 {
		t.Errorf("plain sheet COLINFO records = %d, want 0", got)
	}
	if got := int(binary.LittleEndian.Uint16(rowRecords(t, wideStream)[0][6:8])); got != 800 {
		t.Errorf("wide sheet row 0 height = %d, want 800 (40 points)", got)
	}
	if got := int(binary.LittleEndian.Uint16(rowRecords(t, plainStream)[0][6:8])); got != DefaultRowHeight {
		t.Errorf("plain sheet row 0 height = %d, want %d", got, DefaultRowHeight)
	}
}
