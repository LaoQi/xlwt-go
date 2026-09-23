package xlwt

import (
	"bytes"
	"fmt"
	"testing"
)

// TestWorksheet_WriteOrderDoesNotChangeOutput guards against the porting
// behaviour where the SST index of a string was decided by the order in which
// Worksheet.Write was called. Two workbooks holding the same cells produced
// different bytes depending on the call order, which made it impossible to
// compare a product byte for byte (for example when the caller fills the sheet
// from a Go map).
func TestWorksheet_WriteOrderDoesNotChangeOutput(t *testing.T) {
	build := func(order [][2]int) []byte {
		wb := NewWorkbook()
		ws, err := wb.AddSheet("Sheet1")
		if err != nil {
			t.Fatal(err)
		}
		for _, rc := range order {
			// Repeated values keep the first string occurrence away from the
			// first cell in row/column order.
			if err := ws.Write(rc[0], rc[1], fmt.Sprintf("value_%d", (rc[0]+rc[1])%3)); err != nil {
				t.Fatal(err)
			}
		}
		return wb.GetBiffData()
	}

	var forward, backward [][2]int
	for r := 0; r < 12; r++ {
		for c := 0; c < 12; c++ {
			forward = append(forward, [2]int{r, c})
		}
	}
	for i := len(forward) - 1; i >= 0; i-- {
		backward = append(backward, forward[i])
	}

	a, b := build(forward), build(backward)
	if !bytes.Equal(a, b) {
		t.Errorf("write order changed the output: %d cells compare unequal", cellCount(t, a))
	}
	if got := cellCount(t, a); got != len(forward) {
		t.Fatalf("parsed %d cells, want %d", got, len(forward))
	}

	// Both must still decode to the same content.
	ac := sheetCells(t, a, sheetOffset(t, a))
	bc := sheetCells(t, b, sheetOffset(t, b))
	if len(ac) != len(forward) || len(bc) != len(forward) {
		t.Fatalf("cells = %d and %d, want %d", len(ac), len(bc), len(forward))
	}
	for key, value := range ac {
		if bc[key] != value {
			t.Errorf("cell %v = %q and %q", key, value, bc[key])
		}
	}
}

// TestWorksheet_SortedWriteOrderUnchanged pins the index layout to the position
// of the first cell that references each string: writing cells in row/column
// order must produce the layout an unsorted table would have produced, so a
// caller that sorts its input keeps the historical bytes.
func TestWorksheet_SortedWriteOrderUnchanged(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	// "zzz" is written first, so it keeps index 0 even though "aaa" sorts
	// before it.
	for r, value := range []string{"zzz", "aaa", "mmm"} {
		if err := ws.Write(r, 0, value); err != nil {
			t.Fatal(err)
		}
	}

	want := []string{"zzz", "aaa", "mmm"}
	if sst := wb.SST; !equalStrings(sst.StrList, want) {
		t.Errorf("string table = %q, want %q", sst.StrList, want)
	}
	for _, idx := range []int{0, 1, 2} {
		if got := ws.Grid[uint32(idx<<16)].SSTIdx; got != idx {
			t.Errorf("cell row %d has index %d, want %d", idx, got, idx)
		}
	}

	// Repeated serialisation must not shuffle the table again.
	first := wb.GetBiffData()
	if second := wb.GetBiffData(); !bytes.Equal(first, second) {
		t.Error("two consecutive serialisations produced different bytes")
	}
	if sst := wb.SST; !equalStrings(sst.StrList, want) {
		t.Errorf("string table changed after serialising: %q", sst.StrList)
	}
}

func equalStrings(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func sheetOffset(t *testing.T, stream []byte) int {
	t.Helper()
	sheets := boundSheets(t, stream)
	if len(sheets) == 0 {
		t.Fatal("no BOUNDSHEET record")
	}
	return int(sheets[0].Pos)
}

func cellCount(t *testing.T, stream []byte) int {
	t.Helper()
	return len(sheetCells(t, stream, sheetOffset(t, stream)))
}
