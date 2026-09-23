package xlwt

import (
	"encoding/binary"
	"strings"
	"testing"
	"unicode/utf16"
)

// parseBiffRecords walks a BIFF stream and returns all top-level records.
func parseBiffRecords(t *testing.T, stream []byte) [][2]interface{} {
	t.Helper()
	var out [][2]interface{}
	pos := 0
	for pos+4 <= len(stream) {
		id := binary.LittleEndian.Uint16(stream[pos:])
		size := int(binary.LittleEndian.Uint16(stream[pos+2:]))
		if pos+4+size > len(stream) {
			t.Fatalf("record 0x%04X at %d claims %d bytes, past end of stream (%d)", id, pos, size, len(stream))
		}
		out = append(out, [2]interface{}{id, stream[pos+4 : pos+4+size]})
		pos += 4 + size
	}
	if pos != len(stream) {
		t.Fatalf("trailing %d bytes in stream", len(stream)-pos)
	}
	return out
}

// sstPayloads collects the SST record payload plus every following CONTINUE
// payload, which together form the logical string table.
func sstPayloads(t *testing.T, stream []byte) [][]byte {
	t.Helper()
	records := parseBiffRecords(t, stream)
	var payloads [][]byte
	found := false
	for _, rec := range records {
		id := rec[0].(uint16)
		data := rec[1].([]byte)
		switch {
		case id == SST_ID:
			if found {
				t.Fatal("more than one SST record")
			}
			found = true
			payloads = append(payloads, data)
		case id == CONTINUE_ID && found:
			payloads = append(payloads, data)
		case found && id != CONTINUE_ID:
			return payloads // SST chain ended
		}
	}
	if !found {
		t.Fatal("no SST record found")
	}
	return payloads
}

// parseSSTStrings decodes the shared string table, including strings that were
// split across CONTINUE records. When a string spans a record boundary the
// following record begins with a repeated option flags byte, which must be
// skipped while reading character data.
func parseSSTStrings(t *testing.T, payloads [][]byte) (total, unique uint32, strings []string) {
	t.Helper()

	var buf []byte
	boundaries := map[int]bool{} // offsets where a CONTINUE payload starts
	for i, p := range payloads {
		if i > 0 {
			boundaries[len(buf)] = true
		}
		buf = append(buf, p...)
	}
	if len(buf) < 8 {
		t.Fatalf("SST payload too short: %d bytes", len(buf))
	}
	total = binary.LittleEndian.Uint32(buf[0:4])
	unique = binary.LittleEndian.Uint32(buf[4:8])

	pos := 8
	for i := uint32(0); i < unique; i++ {
		if pos+3 > len(buf) {
			t.Fatalf("string %d: header past end of SST", i)
		}
		cch := int(binary.LittleEndian.Uint16(buf[pos:]))
		flags := buf[pos+2]
		pos += 3

		if flags&0x01 == 0 {
			t.Fatalf("string %d: unexpected compressed (8-bit) encoding", i)
		}
		units := make([]uint16, 0, cch)
		for j := 0; j < cch; j++ {
			if boundaries[pos] {
				pos++ // repeated option flags byte at the start of a CONTINUE
			}
			if pos+2 > len(buf) {
				t.Fatalf("string %d: character data past end of SST", i)
			}
			units = append(units, binary.LittleEndian.Uint16(buf[pos:]))
			pos += 2
		}
		strings = append(strings, string(utf16.Decode(units)))
	}
	return total, unique, strings
}

func TestSharedStringTable_RecordLayout(t *testing.T) {
	sst := NewSharedStringTable()
	values := []string{"alpha", "beta", "alpha", "中文"}
	for _, v := range values {
		sst.AddStr(v)
	}

	payloads := sstPayloads(t, sst.GetBiffRecord())
	for i, p := range payloads {
		if len(p) > MaxSSTLength {
			t.Errorf("payload %d is %d bytes, exceeds the %d byte limit", i, len(p), MaxSSTLength)
		}
	}

	total, unique, got := parseSSTStrings(t, payloads)
	if want := uint32(len(values)); total != want {
		t.Errorf("total string references = %d, want %d", total, want)
	}
	if want := uint32(3); unique != want {
		t.Errorf("unique strings = %d, want %d", unique, want)
	}
	want := []string{"alpha", "beta", "中文"}
	if len(got) != len(want) {
		t.Fatalf("decoded %d strings, want %d", len(got), len(want))
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("string %d = %q, want %q", i, got[i], want[i])
		}
	}
}

// TestSharedStringTable_LongStringsUsedToBeDropped guards against the porting
// bug where any string whose packed form reached MaxSSTCellLength was skipped
// while its index stayed assigned, corrupting the table for every reader.
func TestSharedStringTable_LongStringsUsedToBeDropped(t *testing.T) {
	sst := NewSharedStringTable()

	long := strings.Repeat("A", 5000)      // packed: 10003 bytes
	multibyte := strings.Repeat("测", 3000) // packed: 6003 bytes
	astral := strings.Repeat("😀", 2000)    // surrogate pairs
	huge := strings.Repeat("Z", 40000)     // spans several records
	for _, v := range []string{"before", long, multibyte, astral, huge, "after"} {
		sst.AddStr(v)
	}

	raw := sst.GetBiffRecord()
	payloads := sstPayloads(t, raw)
	if len(payloads) < 2 {
		t.Fatal("expected the SST to be split across CONTINUE records")
	}
	for i, p := range payloads {
		if len(p) > MaxSSTLength {
			t.Errorf("payload %d is %d bytes, exceeds the %d byte limit", i, len(p), MaxSSTLength)
		}
	}

	_, unique, got := parseSSTStrings(t, payloads)
	if want := uint32(len(sst.StrList)); unique != want {
		t.Errorf("unique strings = %d, want %d", unique, want)
	}
	if len(got) != len(sst.StrList) {
		t.Fatalf("decoded %d strings, want %d", len(got), len(sst.StrList))
	}
	for i, want := range sst.StrList {
		if got[i] != want {
			t.Errorf("string %d: decoded %d runes, want %d (equal: %v)", i, len([]rune(got[i])), len([]rune(want)), got[i] == want)
		}
	}
}

// TestSharedStringTable_NoCONTINUENeeded covers the single-record path: a table
// small enough to fit into one SST record must not emit any CONTINUE record.
func TestSharedStringTable_NoCONTINUENeeded(t *testing.T) {
	sst := NewSharedStringTable()
	for i := 0; i < 10; i++ {
		sst.AddStr(strings.Repeat("x", i+1))
	}
	payloads := sstPayloads(t, sst.GetBiffRecord())
	if len(payloads) != 1 {
		t.Errorf("got %d records, want a single SST record", len(payloads))
	}
}
