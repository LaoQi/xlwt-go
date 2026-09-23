package xlwt

import (
	"encoding/binary"
	"strconv"
	"strings"
	"testing"
)

const labelSSTID = 0x00FD

// sheetCells parses the worksheet substream that starts at offset pos and
// returns its cells as {row,col} -> shared string.
func sheetCells(t *testing.T, stream []byte, pos int) map[[2]int]string {
	t.Helper()

	// Either the SST record or the whole SST chain starts before the sheets,
	// so decode and concatenate the shared strings once.
	payloads := sstPayloads(t, stream)
	_, _, strings := parseSSTStrings(t, payloads)

	cells := map[[2]int]string{}
	i := pos
	for i+4 <= len(stream) {
		id := binary.LittleEndian.Uint16(stream[i:])
		size := int(binary.LittleEndian.Uint16(stream[i+2:]))
		data := stream[i+4 : i+4+size]
		if id == eofID {
			break
		}
		if id == labelSSTID {
			if len(data) != 10 {
				t.Fatalf("LABELSST payload is %d bytes, want 10", len(data))
			}
			row := int(binary.LittleEndian.Uint16(data[0:2]))
			col := int(binary.LittleEndian.Uint16(data[2:4]))
			idx := int(binary.LittleEndian.Uint32(data[6:10]))
			if idx < 0 || idx >= len(strings) {
				t.Fatalf("cell (%d,%d) references SST index %d, but only %d strings exist", row, col, idx, len(strings))
			}
			cells[[2]int{row, col}] = strings[idx]
		}
		i += 4 + size
	}
	return cells
}

// TestWorksheet_CellsPointAtTheirOwnSheet is the content-level counterpart of
// TestWorkbook_BoundSheetOffsetsUsedToBeIdentical: it reads every worksheet
// through its BOUNDSHEET offset and checks each sheet holds its own cells.
func TestWorksheet_CellsPointAtTheirOwnSheet(t *testing.T) {
	wb := NewWorkbook()
	names := []string{"S1", "S2", "S3"}
	for _, name := range names {
		ws, err := wb.AddSheet(name)
		if err != nil {
			t.Fatal(err)
		}
		if err := ws.Write(0, 0, "value-"+name); err != nil {
			t.Fatal(err)
		}
		if err := ws.Write(1, 1, "second-"+name); err != nil {
			t.Fatal(err)
		}
	}

	stream := wb.GetBiffData()
	sheets := boundSheets(t, stream)
	if len(sheets) != len(names) {
		t.Fatalf("found %d BOUNDSHEET records, want %d", len(sheets), len(names))
	}

	for i, sh := range sheets {
		cells := sheetCells(t, stream, int(sh.Pos))
		want := map[[2]int]string{
			{0, 0}: "value-" + names[i],
			{1, 1}: "second-" + names[i],
		}
		if len(cells) != len(want) {
			t.Fatalf("sheet %q has %d cells, want %d (%v)", sh.Name, len(cells), len(want), cells)
		}
		for k, v := range want {
			if cells[k] != v {
				t.Errorf("sheet %q cell %v = %q, want %q", sh.Name, k, cells[k], v)
			}
		}
	}
}

// TestWorksheet_WriteOverwritesSameCell documents that writing the same cell
// twice keeps the last value, and that the SST still holds both strings.
func TestWorksheet_WriteOverwritesSameCell(t *testing.T) {
	sst := NewSharedStringTable()
	ws := NewWorksheet("Sheet1", sst)
	if err := ws.Write(0, 0, "first"); err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(0, 0, "second"); err != nil {
		t.Fatal(err)
	}

	if got := len(ws.Grid); got != 1 {
		t.Errorf("Grid size = %d, want 1 (same cell written twice)", got)
	}
	cell := ws.Grid[0]
	if want := sst.StrIndexes["second"]; cell.SSTIdx != want {
		t.Errorf("cell SST index = %d, want %d (last write wins)", cell.SSTIdx, want)
	}
}

// TestWorksheet_SparseLayout checks that row/column bookkeeping survives wide
// gaps between written cells.
func TestWorksheet_SparseLayout(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("Sparse")
	if err != nil {
		t.Fatal(err)
	}
	for _, w := range []struct {
		r, c int
		s    string
	}{
		{0, 0, ""},
		{5, 3, "middle"},
		{99, 50, "far corner"},
	} {
		if err := ws.Write(w.r, w.c, w.s); err != nil {
			t.Fatal(err)
		}
	}

	stream := wb.GetBiffData()
	sheets := boundSheets(t, stream)
	cells := sheetCells(t, stream, int(sheets[0].Pos))

	if len(cells) != 3 {
		t.Fatalf("parsed %d cells, want 3: %v", len(cells), cells)
	}
	if cells[[2]int{0, 0}] != "" {
		t.Errorf("cell (0,0) = %q, want empty string", cells[[2]int{0, 0}])
	}
	if got := cells[[2]int{5, 3}]; got != "middle" {
		t.Errorf("cell (5,3) = %q, want %q", got, "middle")
	}
	if got := cells[[2]int{99, 50}]; got != "far corner" {
		t.Errorf("cell (99,50) = %q, want %q", got, "far corner")
	}
}

// TestWorksheet_WriteValidation covers the range and length checks added to
// Write.
func TestWorksheet_WriteValidation(t *testing.T) {
	cases := []struct {
		name    string
		r, c    int
		label   string
		wantErr error
	}{
		{"first cell", 0, 0, "ok", nil},
		{"last cell", MaxRow, MaxCol, "ok", nil},
		{"negative row", -1, 0, "x", ErrRowOutOfRange},
		{"row past limit", MaxRow + 1, 0, "x", ErrRowOutOfRange},
		{"negative column", 0, -1, "x", ErrColOutOfRange},
		{"column past limit", 0, MaxCol + 1, "x", ErrColOutOfRange},
		{"max length string", 0, 0, strings.Repeat("a", MaxStringLength), nil},
		{"string too long", 0, 0, strings.Repeat("a", MaxStringLength+1), ErrStringTooLong},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			ws := NewWorksheet("Sheet1", NewSharedStringTable())
			err := ws.Write(tc.r, tc.c, tc.label)

			if tc.wantErr == nil {
				if err != nil {
					t.Fatalf("Write(%d, %d, %d chars) = %v, want nil", tc.r, tc.c, len(tc.label), err)
				}
				if len(ws.Grid) != 1 {
					t.Errorf("cell was not stored: Grid = %v", ws.Grid)
				}
				return
			}

			if err == nil {
				t.Fatalf("Write(%d, %d, %d chars) = nil, want %v", tc.r, tc.c, len(tc.label), tc.wantErr)
			}
			if !errorsIs(err, tc.wantErr) {
				t.Errorf("Write error = %v, want it to wrap %v", err, tc.wantErr)
			}
			if len(ws.Grid) != 0 || len(ws.RowsIndex) != 0 {
				t.Errorf("worksheet was modified despite the error: Grid=%v RowsIndex=%v", ws.Grid, ws.RowsIndex)
			}
		})
	}
}

// TestWorksheet_SerialisationScalesLinearly guards against the quadratic writer
// that scanned the whole grid once per row. The check is on output identity
// rather than on time: every row must keep its own cells, in column order, no
// matter how the grid is laid out.
func TestWorksheet_SerialisationScalesLinearly(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("Wide")
	if err != nil {
		t.Fatal(err)
	}
	// A tall, sparse sheet: 3000 rows, one cell each, columns widely spread.
	values := map[[2]int]string{}
	for r := 0; r < 3000; r++ {
		c := (r * 7) % MaxCol
		value := "row_" + strconv.Itoa(r)
		values[[2]int{r, c}] = value
		if err := ws.Write(r, c, value); err != nil {
			t.Fatal(err)
		}
	}
	// One row with many cells, to keep the column sorting exercised.
	for c := MaxCol; c >= 0; c-- {
		value := "top_" + strconv.Itoa(c)
		values[[2]int{0, c}] = value
		if err := ws.Write(0, c, value); err != nil {
			t.Fatal(err)
		}
	}

	stream := wb.GetBiffData()
	sheets := boundSheets(t, stream)
	cells := sheetCells(t, stream, int(sheets[0].Pos))
	if len(cells) != len(values) {
		t.Fatalf("parsed %d cells, want %d", len(cells), len(values))
	}
	for key, want := range values {
		if got := cells[key]; got != want {
			t.Errorf("cell %v = %q, want %q", key, got, want)
		}
	}

	// Every row record must be followed by its own cells, in column order.
	records := parseBiffRecords(t, stream)
	for i, rec := range records {
		if rec[0].(uint16) != 0x0208 {
			continue
		}
		row := int(binary.LittleEndian.Uint16(rec[1].([]byte)[0:2]))
		lastCol := -1
		for j := i + 1; j < len(records) && records[j][0].(uint16) == labelSSTID; j++ {
			data := records[j][1].([]byte)
			if got := int(binary.LittleEndian.Uint16(data[0:2])); got != row {
				t.Fatalf("row %d record contains a cell of row %d", row, got)
			}
			col := int(binary.LittleEndian.Uint16(data[2:4]))
			if col <= lastCol {
				t.Fatalf("row %d cells are not ordered: %d after %d", row, col, lastCol)
			}
			lastCol = col
		}
	}
}

// TestWorksheet_StringLengthCountsCharacters checks that the 32767 limit is
// measured in characters, not bytes, so multi-byte text is not rejected early.
func TestWorksheet_StringLengthCountsCharacters(t *testing.T) {
	ws := NewWorksheet("Sheet1", NewSharedStringTable())

	// 20000 CJK characters: 60000 bytes in UTF-8, but only 20000 characters.
	if err := ws.Write(0, 0, strings.Repeat("测", 20000)); err != nil {
		t.Errorf("Write of 20000 CJK characters = %v, want nil", err)
	}
	// Astral characters (emoji) count as well; 32768 of them exceed the limit.
	if err := ws.Write(1, 0, strings.Repeat("😀", MaxStringLength+1)); !errorsIs(err, ErrStringTooLong) {
		t.Errorf("Write of %d emoji = %v, want ErrStringTooLong", MaxStringLength+1, err)
	}
}
