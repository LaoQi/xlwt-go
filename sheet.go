package xlwt

import "sort"

// finalizeSST makes the byte layout of a workbook independent of the order in
// which the cells were written.
//
// Shared string indexes are assigned when a string is first added to the table,
// so they follow the call order of Worksheet.Write, while cells are always
// written sorted by row and column. Two workbooks holding the same cells but
// written in a different order therefore used to produce the same cells and a
// different (although equivalent) string table.
//
// The table is sorted by the position of the first cell that refers to each
// string, walking worksheets in order and cells by row and column. A workbook
// that writes its cells in that same order - the usual case, and what a caller
// iterating a map can restore with one sort - therefore keeps the index layout
// it had before this reordering existed, byte for byte. Any other write order
// yields exactly the same result as well, because the reference order no longer
// depends on the order of the calls.
//
// Cells whose string was overwritten keep their table position, so an
// unreferenced string still occupies an index (both behaviours match the
// Python original).
//
// Every cell is rewritten to its new index. The workbook is serialised as
// [Global] records, then the SST, then the sheets, so the table is final by the
// time any record quoting an index is built.
func (wb *Workbook) finalizeSST() {
	sst := wb.SST
	if sst == nil || len(sst.StrList) == 0 {
		return
	}

	// firstUse records where each string is referenced for the first time.
	type cellPos struct{ sheet, row, col int }
	firstUse := make(map[int]cellPos, len(sst.StrList))
	for sheet, ws := range wb.Worksheets {
		for _, cell := range ws.Grid {
			if cell.SSTIdx < 0 || cell.SSTIdx >= len(sst.StrList) {
				continue
			}
			pos, seen := firstUse[cell.SSTIdx]
			if !seen || pos.sheet > sheet ||
				(pos.sheet == sheet && (pos.row > cell.Row ||
					(pos.row == cell.Row && pos.col > cell.Col))) {
				firstUse[cell.SSTIdx] = cellPos{sheet, cell.Row, cell.Col}
			}
		}
	}

	// Strings that no cell refers to (for example after a cell was overwritten)
	// keep their relative order and are placed before the referenced ones.
	order := make([]int, len(sst.StrList))
	for i := range order {
		order[i] = i
	}
	sort.SliceStable(order, func(a, b int) bool {
		pa, foundA := firstUse[order[a]]
		pb, foundB := firstUse[order[b]]
		if foundA != foundB {
			return foundA // unreferenced strings come first
		}
		if !foundA {
			return order[a] < order[b]
		}
		if pa.sheet != pb.sheet {
			return pa.sheet < pb.sheet
		}
		if pa.row != pb.row {
			return pa.row < pb.row
		}
		return pa.col < pb.col
	})

	sorted := make([]string, len(sst.StrList))
	remap := make([]int, len(sst.StrList))
	for newIdx, oldIdx := range order {
		sorted[newIdx] = sst.StrList[oldIdx]
		remap[oldIdx] = newIdx
	}
	sst.StrList = sorted
	for value, oldIdx := range sst.StrIndexes {
		sst.StrIndexes[value] = remap[oldIdx]
	}
	for _, ws := range wb.Worksheets {
		ws.remapSSTIndexes(remap)
	}
}
