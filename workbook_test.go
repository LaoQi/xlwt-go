package xlwt

import (
	"bytes"
	"encoding/binary"
	"strings"
	"testing"
)

const boundSheetID = 0x0085
const bofID = 0x0809
const eofID = 0x000A

// decodeBoundSheet returns the stream position, sheet name and payload of a
// BOUNDSHEET record.
func decodeBoundSheet(t *testing.T, data []byte) (pos uint32, name string) {
	t.Helper()
	if len(data) < 8 {
		t.Fatalf("BOUNDSHEET payload too short: %d bytes", len(data))
	}
	pos = binary.LittleEndian.Uint32(data[0:4])
	u8len := int(data[6])
	flags := data[7]

	var units []byte
	if flags&0x01 == 0 {
		if len(data) < 8+u8len {
			t.Fatalf("BOUNDSHEET 8-bit name length %d exceeds payload %d", u8len, len(data))
		}
		units = data[8 : 8+u8len]
		// Compressed format: every byte is one character whose code point is
		// the byte value itself (latin-1), not a UTF-8 sequence.
		runes := make([]rune, 0, u8len)
		for _, b := range units {
			runes = append(runes, rune(b))
		}
		name = string(runes)
	} else {
		if len(data) < 8+2*u8len {
			t.Fatalf("BOUNDSHEET 16-bit name length %d exceeds payload %d", u8len, len(data))
		}
		units = data[8 : 8+2*u8len]
		runes := make([]rune, 0, u8len)
		for i := 0; i+1 < len(units); i += 2 {
			runes = append(runes, rune(binary.LittleEndian.Uint16(units[i:])))
		}
		name = string(runes)
	}
	return pos, name
}

type boundSheetInfo struct {
	Pos  uint32
	Name string
}

func boundSheets(t *testing.T, stream []byte) []boundSheetInfo {
	t.Helper()
	var out []boundSheetInfo
	for _, rec := range parseBiffRecords(t, stream) {
		if rec[0].(uint16) == boundSheetID {
			pos, name := decodeBoundSheet(t, rec[1].([]byte))
			out = append(out, boundSheetInfo{pos, name})
		}
	}
	return out
}

// TestWorkbook_BoundSheetOffsetsUsedToBeIdentical guards against the porting bug
// where every BOUNDSHEET record carried the same stream position, making every
// reader show the first worksheet's contents for all sheets.
func TestWorkbook_BoundSheetOffsetsUsedToBeIdentical(t *testing.T) {
	wb := NewWorkbook()
	names := []string{"S1", "S2", "S3", "S4"}
	for _, name := range names {
		ws, err := wb.AddSheet(name)
		if err != nil {
			t.Fatal(err)
		}
		if err := ws.Write(0, 0, "value-for-"+name); err != nil {
			t.Fatal(err)
		}
		// Make the sheets different lengths so strides differ.
		for i := 1; i <= len(name); i++ {
			if err := ws.Write(i, 0, strings.Repeat(name, i)); err != nil {
				t.Fatal(err)
			}
		}
	}

	stream := wb.GetBiffData()
	sheets := boundSheets(t, stream)
	if len(sheets) != len(names) {
		t.Fatalf("found %d BOUNDSHEET records, want %d", len(sheets), len(names))
	}

	for i, sh := range sheets {
		if sh.Name != names[i] {
			t.Errorf("BOUNDSHEET %d name = %q, want %q", i, sh.Name, names[i])
		}
		if i > 0 && sh.Pos == sheets[i-1].Pos {
			t.Errorf("BOUNDSHEET %d and %d share stream position %d; offsets are not advancing", i-1, i, sh.Pos)
		}
	}

	// Each recorded position must point at a worksheet BOF stream.
	for i, sh := range sheets {
		if int(sh.Pos)+4 > len(stream) {
			t.Fatalf("sheet %d: stream position %d past end of stream (%d)", i, sh.Pos, len(stream))
		}
		id := binary.LittleEndian.Uint16(stream[sh.Pos:])
		if id != bofID {
			t.Errorf("sheet %d: stream position %d points at record 0x%04X, want worksheet BOF 0x%04X", i, sh.Pos, id, bofID)
		}
	}
}

// TestWorkbook_BoundSheetOffsetsMatchSheetLengths checks that the offset stride
// equals the previous worksheet's BIFF length, which is what upstream xlwt does.
func TestWorkbook_BoundSheetOffsetsMatchSheetLengths(t *testing.T) {
	wb := NewWorkbook()
	if _, err := wb.AddSheet("A"); err != nil {
		t.Fatal(err)
	}
	wsB, err := wb.AddSheet("B")
	if err != nil {
		t.Fatal(err)
	}
	if err := wsB.Write(0, 0, "only B has a cell"); err != nil {
		t.Fatal(err)
	}

	stream := wb.GetBiffData()
	sheets := boundSheets(t, stream)
	if len(sheets) != 2 {
		t.Fatalf("found %d BOUNDSHEET records, want 2", len(sheets))
	}
	if sheets[0].Pos == sheets[1].Pos {
		t.Fatalf("both sheets point at %d", sheets[0].Pos)
	}

	// Locate sheet B's BOF, and confirm sheet A's data sits between them.
	bofB := int(sheets[1].Pos)
	if bofB+4 > len(stream) || binary.LittleEndian.Uint16(stream[bofB:]) != bofID {
		t.Fatalf("sheet B offset %d does not point at a BOF record", bofB)
	}
	// Sheet A must span [sheets[0].Pos, bofB) and end with an EOF record.
	aStart := int(sheets[0].Pos)
	if aStart >= bofB {
		t.Fatalf("sheet A range [%d,%d) is empty", aStart, bofB)
	}
	if binary.LittleEndian.Uint16(stream[aStart:]) != bofID {
		t.Errorf("sheet A offset %d does not point at a BOF record", aStart)
	}
	eofPos := bofB - 4 // EOF record is the last record of a worksheet substream
	if binary.LittleEndian.Uint16(stream[eofPos:]) != eofID {
		t.Errorf("record before sheet B is 0x%04X, want worksheet EOF 0x%04X", binary.LittleEndian.Uint16(stream[eofPos:]), eofID)
	}
}

// TestWorkbook_SaveMultiSheetRoundTrip writes a multi-sheet workbook and checks
// the stream shape: one globals BOF, one BOUNDSHEET per sheet, and one BOF per
// worksheet.
func TestWorkbook_SaveMultiSheetRoundTrip(t *testing.T) {
	wb := NewWorkbook()
	for i := 0; i < 5; i++ {
		ws, err := wb.AddSheet(string(rune('A' + i)))
		if err != nil {
			t.Fatal(err)
		}
		if err := ws.Write(0, 0, string(rune('a'+i))); err != nil {
			t.Fatal(err)
		}
	}

	var buf bytes.Buffer
	if err := wb.Save(&buf); err != nil {
		t.Fatal(err)
	}
	if !bytes.HasPrefix(buf.Bytes(), ole2Signature) {
		t.Fatalf("output is not an OLE2 compound document")
	}

	stream := wb.GetBiffData()
	records := parseBiffRecords(t, stream)
	var bofs, boundSheetsCount int
	for _, rec := range records {
		switch rec[0].(uint16) {
		case bofID:
			bofs++
		case boundSheetID:
			boundSheetsCount++
		}
	}
	if bofs != 6 { // 1 workbook globals + 5 worksheets
		t.Errorf("found %d BOF records, want 6", bofs)
	}
	if boundSheetsCount != 5 {
		t.Errorf("found %d BOUNDSHEET records, want 5", boundSheetsCount)
	}
}

// TestWorkbook_AddSheetValidation covers the worksheet name rules.
func TestWorkbook_AddSheetValidation(t *testing.T) {
	valid := []string{"Sheet1", "a", strings.Repeat("x", MaxSheetNameLength), "中文表名", "with space"}
	for _, name := range valid {
		wb := NewWorkbook()
		if _, err := wb.AddSheet(name); err != nil {
			t.Errorf("AddSheet(%q) = %v, want nil", name, err)
		}
	}

	invalid := []string{"", strings.Repeat("x", MaxSheetNameLength+1), "'leading", "a/b", "a\\b", "a?b", "a*b", "a[b", "a]b", "a:b", "a\x00b"}
	for _, name := range invalid {
		wb := NewWorkbook()
		_, err := wb.AddSheet(name)
		if err == nil {
			t.Errorf("AddSheet(%q) = nil, want ErrInvalidSheetName", name)
			continue
		}
		if !errorsIs(err, ErrInvalidSheetName) {
			t.Errorf("AddSheet(%q) = %v, want it to wrap ErrInvalidSheetName", name, err)
		}
		if len(wb.Worksheets) != 0 {
			t.Errorf("AddSheet(%q) added a sheet despite the error", name)
		}
	}
}

// TestWorkbook_AddSheetRejectsDuplicates checks case-insensitive duplicate name
// detection, which is what Excel and upstream xlwt enforce.
func TestWorkbook_AddSheetRejectsDuplicates(t *testing.T) {
	wb := NewWorkbook()
	if _, err := wb.AddSheet("Data"); err != nil {
		t.Fatal(err)
	}

	for _, dup := range []string{"Data", "data", "DATA"} {
		_, err := wb.AddSheet(dup)
		if err == nil {
			t.Errorf("AddSheet(%q) = nil, want ErrDuplicateSheetName", dup)
			continue
		}
		if !errorsIs(err, ErrDuplicateSheetName) {
			t.Errorf("AddSheet(%q) = %v, want it to wrap ErrDuplicateSheetName", dup, err)
		}
	}
	if len(wb.Worksheets) != 1 {
		t.Errorf("workbook has %d sheets, want 1", len(wb.Worksheets))
	}
}

// TestWorkbook_NonLatinSheetNameUsedToBeMangled guards against the porting bug
// where sheet names were always written in the compressed 8-bit format, which
// corrupted any name outside latin-1 (for example CJK).
func TestWorkbook_NonLatinSheetNameUsedToBeMangled(t *testing.T) {
	wb := NewWorkbook()
	names := []string{"ASCII", "中文表", "日本語", "Кириллица", "Ünïcödé"}
	for _, name := range names {
		if _, err := wb.AddSheet(name); err != nil {
			t.Fatal(err)
		}
	}

	stream := wb.GetBiffData()
	sheets := boundSheets(t, stream)
	if len(sheets) != len(names) {
		t.Fatalf("found %d BOUNDSHEET records, want %d", len(sheets), len(names))
	}
	for i, sh := range sheets {
		if sh.Name != names[i] {
			t.Errorf("BOUNDSHEET %d name = %q, want %q", i, sh.Name, names[i])
		}
	}
}
