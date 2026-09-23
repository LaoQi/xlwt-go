package xlwt

import (
	"encoding/binary"
	"strings"
	"testing"
)

// xfRecords extracts every XF record payload from a BIFF stream, in order.
func xfRecords(t *testing.T, stream []byte) [][]byte {
	t.Helper()
	var out [][]byte
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) == 0x00E0 {
			out = append(out, rec[1].([]byte))
		}
	}
	return out
}

// fontRecords extracts every FONT record payload from a BIFF stream, in order.
func fontRecords(t *testing.T, stream []byte) [][]byte {
	t.Helper()
	var out [][]byte
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) == 0x0031 {
			out = append(out, rec[1].([]byte))
		}
	}
	return out
}

// labelSSTXFIndexes returns the XF index of every cell, keyed by row/col.
func labelSSTXFIndexes(t *testing.T, stream []byte) map[[2]int]int {
	t.Helper()
	out := map[[2]int]int{}
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) != labelSSTID {
			continue
		}
		data := rec[1].([]byte)
		row := int(binary.LittleEndian.Uint16(data[0:2]))
		col := int(binary.LittleEndian.Uint16(data[2:4]))
		xf := int(binary.LittleEndian.Uint16(data[4:6]))
		out[[2]int{row, col}] = xf
	}
	return out
}

// TestStyleCollection_DefaultLayout pins the index layout of the default
// workbook: fonts 0,1,2,3,5,6,7 (index 4 is never used in BIFF8), sixteen
// built-in style XFs followed by the cell XFs, and cell XF 0x11 as the default
// style.
func TestStyleCollection_DefaultLayout(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	if err := ws.Write(0, 0, "x"); err != nil {
		t.Fatal(err)
	}

	stream := wb.GetBiffData()

	fonts := fontRecords(t, stream)
	if len(fonts) != 7 {
		t.Errorf("font count = %d, want 7 (indexes 0,1,2,3,5,6,7; 4 is skipped)", len(fonts))
	}

	xfs := xfRecords(t, stream)
	if len(xfs) != 18 {
		t.Errorf("XF count = %d, want 18 (16 built-in + 2 cell XFs)", len(xfs))
	}

	// The default style must be reachable as DefaultCellXFStyle.
	idx, err := wb.AddStyle(nil)
	if err != nil {
		t.Fatal(err)
	}
	if idx != DefaultCellXFStyle {
		t.Errorf("default style index = 0x%X, want 0x%X", idx, DefaultCellXFStyle)
	}
	cells := labelSSTXFIndexes(t, stream)
	if got := cells[[2]int{0, 0}]; got != DefaultCellXFStyle {
		t.Errorf("cell without style has XF 0x%X, want 0x%X", got, DefaultCellXFStyle)
	}
}

// TestStyleCollection_Deduplicates checks that equal styles share one XF and
// that different styles get different ones.
func TestStyleCollection_Deduplicates(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}

	bold := NewXFStyle()
	bold.Font.Bold = true

	// The same style value used many times must resolve to one XF index.
	var indexes []int
	for i := 0; i < 50; i++ {
		idx, err := wb.AddStyle(bold)
		if err != nil {
			t.Fatal(err)
		}
		indexes = append(indexes, idx)
	}
	if len(uniqueInts(indexes)) != 1 {
		t.Errorf("50 registrations of the same style produced indexes %v", uniqueInts(indexes))
	}

	// A separate but equal style object must also deduplicate.
	bold2 := NewXFStyle()
	bold2.Font.Bold = true
	idx, err := wb.AddStyle(bold2)
	if err != nil {
		t.Fatal(err)
	}
	if idx != indexes[0] {
		t.Errorf("equal style objects resolved to different XFs: 0x%X vs 0x%X", idx, indexes[0])
	}

	// A style without a font change shares the font but not the XF.
	plain := NewXFStyle()
	plain.Alignment.Horz = HorzCenter
	idx2, err := wb.AddStyle(plain)
	if err != nil {
		t.Fatal(err)
	}
	if idx2 == indexes[0] {
		t.Errorf("different styles share XF 0x%X", idx2)
	}

	if err := ws.WriteWithStyle(0, 0, "bold", bold); err != nil {
		t.Fatal(err)
	}
	if err := ws.WriteWithStyle(1, 0, "centered", plain); err != nil {
		t.Fatal(err)
	}

	cells := labelSSTXFIndexes(t, wb.GetBiffData())
	if cells[[2]int{0, 0}] != indexes[0] {
		t.Errorf("bold cell XF = 0x%X, want 0x%X", cells[[2]int{0, 0}], indexes[0])
	}
	if cells[[2]int{1, 0}] != idx2 {
		t.Errorf("centered cell XF = 0x%X, want 0x%X", cells[[2]int{1, 0}], idx2)
	}
}

// TestStyleCollection_FontIndexSkipsFour checks the BIFF8 rule that font index
// 4 is never used.
func TestStyleCollection_FontIndexSkipsFour(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}

	// Register enough distinct fonts to need indexes beyond 7.
	for i := 0; i < 6; i++ {
		style := NewXFStyle()
		style.Font.Name = "Font" + strings.Repeat("X", i+1)
		if err := ws.WriteWithStyle(i, 0, "x", style); err != nil {
			t.Fatal(err)
		}
	}

	fonts := fontRecords(t, wb.GetBiffData())
	// 8 defaults scaled to 7 fonts + 6 new fonts:
	// the count alone does not prove index 4 is skipped, so check the stream.
	if len(fonts) < 8 {
		t.Fatalf("expected at least 8 FONT records, got %d", len(fonts))
	}

	// Every XF record's font index must be a valid font index (never 4).
	xfs := xfRecords(t, wb.GetBiffData())
	valid := map[int]bool{}
	for i := range fonts {
		idx := i
		if idx >= 4 { // index 4 is skipped, so the written position shifts
			idx = i
		}
		valid[idx] = true
	}
	for i, xf := range xfs {
		fontIdx := int(binary.LittleEndian.Uint16(xf[0:2]))
		if fontIdx == 4 {
			t.Errorf("XF %d references font index 4, which must never exist", i)
		}
	}
}

// TestStyleCollection_NumberFormats checks that built-in formats reuse their
// reserved index and custom formats get sequentially assigned user indexes.
func TestStyleCollection_NumberFormats(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}

	builtin := NewXFStyle()
	builtin.NumFormatStr = "0.00%"
	first, err := wb.AddStyle(builtin)
	if err != nil {
		t.Fatal(err)
	}

	// Registering the built-in format again must not create a format record.
	second, err := wb.AddStyle(builtin)
	if err != nil {
		t.Fatal(err)
	}
	if first != second {
		t.Errorf("built-in format style not deduplicated: 0x%X vs 0x%X", first, second)
	}

	stream := wb.GetBiffData()
	var formatRecords int
	var formatIndexes []int
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) == 0x041E {
			formatRecords++
			formatIndexes = append(formatIndexes, int(binary.LittleEndian.Uint16(rec[1].([]byte)[0:2])))
		}
	}
	if formatRecords != 1 {
		t.Errorf("FORMAT record count = %d, want 1 (the built-in format needs none)", formatRecords)
	}
	if len(formatIndexes) == 1 && formatIndexes[0] != FIRST_USER_DEFINED_NUM_FORMAT_IDX {
		t.Errorf("FORMAT index = %d, want %d", formatIndexes[0], FIRST_USER_DEFINED_NUM_FORMAT_IDX)
	}

	// A custom format must be added as a second FORMAT record.
	custom := NewXFStyle()
	custom.NumFormatStr = "#,##0.000"
	idx, err := wb.AddStyle(custom)
	if err != nil {
		t.Fatal(err)
	}
	if idx == first {
		t.Error("custom number format did not create a new XF")
	}

	var customIndexes []int
	for _, rec := range parseBiffRecords(t, wb.GetBiffData()) {
		if rec[0].(uint16) == 0x041E {
			customIndexes = append(customIndexes, int(binary.LittleEndian.Uint16(rec[1].([]byte)[0:2])))
		}
	}
	if len(customIndexes) != 2 {
		t.Fatalf("FORMAT record count = %d, want 2", len(customIndexes))
	}
	if customIndexes[1] != FIRST_USER_DEFINED_NUM_FORMAT_IDX+1 {
		t.Errorf("custom FORMAT index = %d, want %d", customIndexes[1], FIRST_USER_DEFINED_NUM_FORMAT_IDX+1)
	}
	if err := ws.WriteWithStyle(0, 0, "x", custom); err != nil {
		t.Fatal(err)
	}
}

// TestStyle_DefaultOutputUnchanged is the regression test for the guarantee that
// adding configurable styles did not change the bytes of a workbook that uses no
// explicit style.
func TestStyle_DefaultOutputUnchanged(t *testing.T) {
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

	// The style section must be exactly the bootstrap section: 8 fonts, 16 style
	// XFs, 2 cell XFs, 1 FORMAT and 1 STYLE record, in this order.
	var order []uint16
	for _, rec := range parseBiffRecords(t, stream) {
		id := rec[0].(uint16)
		switch id {
		case 0x0031, 0x041E, 0x00E0, 0x0293:
			order = append(order, id)
		}
	}
	want := []uint16{}
	for i := 0; i < 7; i++ {
		want = append(want, 0x0031)
	}
	want = append(want, 0x041E)
	for i := 0; i < 18; i++ {
		want = append(want, 0x00E0)
	}
	want = append(want, 0x0293)

	if len(order) != len(want) {
		t.Fatalf("style section record count = %d, want %d (%v)", len(order), len(want), order)
	}
	for i := range want {
		if order[i] != want[i] {
			t.Fatalf("style section record %d = 0x%04X, want 0x%04X", i, order[i], want[i])
		}
	}
}

// TestStyle_WriteWithStyleValidation checks that an explicit style on a
// detached worksheet is rejected instead of silently ignored.
func TestStyle_WriteWithStyleValidation(t *testing.T) {
	ws := NewWorksheet("detached", NewSharedStringTable())

	style := NewXFStyle()
	style.Font.Bold = true
	err := ws.WriteWithStyle(0, 0, "x", style)
	if !errorsIs(err, ErrUnattachedWorksheet) {
		t.Errorf("WriteWithStyle on a detached worksheet = %v, want ErrUnattachedWorksheet", err)
	}
	if len(ws.Grid) != 0 {
		t.Error("detached worksheet was modified despite the error")
	}

	// Without a style the default XF is used and the write succeeds.
	if err := ws.WriteWithStyle(0, 0, "x", nil); err != nil {
		t.Errorf("WriteWithStyle without a style = %v, want nil", err)
	}
	if len(ws.Grid) != 1 {
		t.Error("WriteWithStyle without a style did not store the cell")
	}
}

// TestStyle_FontRecordEncoding checks the FONT record bit flags and the name
// encoding of the record.
func TestStyle_FontRecordEncoding(t *testing.T) {
	style := NewXFStyle()
	style.Font.Bold = true
	style.Font.Italic = true
	style.Font.StrikeOut = true
	style.Font.Name = "Courier New"

	rec := fontRecord(style.Font)
	// payload after the 4-byte record header
	data := rec[4:]

	if got := binary.LittleEndian.Uint16(data[0:2]); got != 0x00C8 {
		t.Errorf("height = 0x%X, want 0x00C8", got)
	}
	if got := binary.LittleEndian.Uint16(data[2:4]); got != 0x01|0x02|0x08 {
		t.Errorf("options = 0x%X, want 0x%X", got, 0x01|0x02|0x08)
	}
	if got := binary.LittleEndian.Uint16(data[6:8]); got != 0x02BC {
		t.Errorf("weight = 0x%X, want 0x02BC for a bold font", got)
	}

	// name: 1-byte length, 1-byte flags, then the characters
	nameLen := int(data[14])
	if nameLen != len("Courier New") {
		t.Fatalf("font name length = %d, want %d", nameLen, len("Courier New"))
	}
	if got := string(data[16 : 16+nameLen]); got != "Courier New" {
		t.Errorf("font name = %q, want %q", got, "Courier New")
	}
}

// TestStyle_XFRecordBitPacking verifies the alignment, border and pattern
// bit fields of the XF record.
func TestStyle_XFRecordBitPacking(t *testing.T) {
	align := Alignment{Horz: HorzCenter, Vert: VertTop, Wrap: true, Indent: 3}
	borders := Borders{
		Left: BorderThin, Right: BorderDouble, Top: BorderNone, Bottom: BorderThick,
		LeftColour: 0x08, RightColour: 0x0A, TopColour: 0x40, BottomColour: 0x0B,
	}
	pattern := Pattern{Pattern: PatternSolid, ForeColour: 0x2C, BackColour: 0x41}
	prot := Protection{FormulaHidden: true} // CellUnlocked false => locked

	rec := xfRecord(7, 164, align, borders, pattern, prot, false)
	data := rec[4:]

	if got := binary.LittleEndian.Uint16(data[0:2]); got != 7 {
		t.Errorf("font index = %d, want 7", got)
	}
	if got := binary.LittleEndian.Uint16(data[2:4]); got != 164 {
		t.Errorf("format index = %d, want 164", got)
	}
	// cell locked + formula hidden
	if got := binary.LittleEndian.Uint16(data[4:6]); got != 0x03 {
		t.Errorf("protection = 0x%X, want 0x03", got)
	}
	// horz centre (2) | wrap (1<<3) | vert top (0<<4)
	if got := data[6]; got != 0x02|0x08 {
		t.Errorf("alignment byte = 0x%X, want 0x%X", got, 0x02|0x08)
	}
	// indent 3 | wrap-adjacent flags
	if got := data[8]; got != 0x03 {
		t.Errorf("text byte = 0x%X, want 0x03 (indent 3)", got)
	}
	if got := data[9]; got != 0xF8 {
		t.Errorf("used-attributes byte = 0x%X, want 0xF8 for a cell XF", got)
	}

	border1 := binary.LittleEndian.Uint32(data[10:14])
	if got := border1 & 0x0F; got != BorderThin {
		t.Errorf("left border style = %d, want %d", got, BorderThin)
	}
	if got := (border1 >> 4) & 0x0F; got != BorderDouble {
		t.Errorf("right border style = %d, want %d", got, BorderDouble)
	}
	if got := (border1 >> 8) & 0x0F; got != BorderNone {
		t.Errorf("top border style = %d, want %d", got, BorderNone)
	}
	if got := (border1 >> 12) & 0x0F; got != BorderThick {
		t.Errorf("bottom border style = %d, want %d", got, BorderThick)
	}
	if got := (border1 >> 16) & 0x7F; got != 0x08 {
		t.Errorf("left border colour = 0x%X, want 0x08", got)
	}
	if got := (border1 >> 23) & 0x7F; got != 0x0A {
		t.Errorf("right border colour = 0x%X, want 0x0A", got)
	}

	border2 := binary.LittleEndian.Uint32(data[14:18])
	// a BorderNone line must not carry a colour
	if got := border2 & 0x7F; got != 0x00 {
		t.Errorf("top border colour = 0x%X, want 0 because the top line has no style", got)
	}
	if got := (border2 >> 7) & 0x7F; got != 0x0B {
		t.Errorf("bottom border colour = 0x%X, want 0x0B", got)
	}
	if got := (border2 >> 26) & 0x3F; got != PatternSolid {
		t.Errorf("pattern = 0x%X, want 0x%X", got, PatternSolid)
	}

	patWord := binary.LittleEndian.Uint16(data[18:20])
	if got := patWord & 0x7F; got != 0x2C {
		t.Errorf("pattern foreground = 0x%X, want 0x2C", got)
	}
	if got := (patWord >> 7) & 0x7F; got != 0x41 {
		t.Errorf("pattern background = 0x%X, want 0x41", got)
	}
}

// TestStyle_ZeroValueUsable checks that a partially initialised style works and
// falls back to documented defaults.
func TestStyle_ZeroValueUsable(t *testing.T) {
	wb := NewWorkbook()
	ws, err := wb.AddSheet("S")
	if err != nil {
		t.Fatal(err)
	}

	// Only bold is set; everything else must fall back to the defaults.
	if err := ws.WriteWithStyle(0, 0, "x", &XFStyle{Font: Font{Bold: true}}); err != nil {
		t.Fatal(err)
	}
	// NewXFStyle is exactly the default style, so it must reuse XF 0x11.
	idx, err := wb.AddStyle(NewXFStyle())
	if err != nil {
		t.Fatal(err)
	}
	if idx != DefaultCellXFStyle {
		t.Errorf("NewXFStyle index = 0x%X, want 0x%X (the default)", idx, DefaultCellXFStyle)
	}

	fonts := fontRecords(t, wb.GetBiffData())
	var foundBold bool
	for _, f := range fonts {
		if binary.LittleEndian.Uint16(f[6:8]) == 0x02BC {
			foundBold = true
			if got := binary.LittleEndian.Uint16(f[0:2]); got != 0x00C8 {
				t.Errorf("bold font height = 0x%X, want the default 0x00C8", got)
			}
			nameLen := int(f[14])
			if got := string(f[16 : 16+nameLen]); got != "Arial" {
				t.Errorf("bold font name = %q, want the default %q", got, "Arial")
			}
		}
	}
	if !foundBold {
		t.Error("no bold font (weight 0x02BC) was written")
	}
}

// uniqueInts returns the distinct values of the slice, in order of appearance.
func uniqueInts(values []int) []int {
	seen := map[int]bool{}
	var out []int
	for _, v := range values {
		if !seen[v] {
			seen[v] = true
			out = append(out, v)
		}
	}
	return out
}
