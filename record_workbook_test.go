package xlwt

import (
	"bytes"
	"strconv"
	"strings"
	"testing"
	"unicode/utf8"
)

// TestWorkbook_OwnerLongNameUsedToPanic guards against the panic raised by a
// negative padding length when the owner name did not fit into the fixed size
// WRITEACCESS record.
func TestWorkbook_OwnerLongNameUsedToPanic(t *testing.T) {
	for _, n := range []int{0, 1, 47, 48, 49, 112, 113, 1000} {
		t.Run("owner_"+strconv.Itoa(n), func(t *testing.T) {
			wb := NewWorkbook()
			wb.Owner = strings.Repeat("x", n)
			ws, err := wb.AddSheet("S")
			if err != nil {
				t.Fatal(err)
			}
			if err := ws.Write(0, 0, "a"); err != nil {
				t.Fatal(err)
			}

			rec := WriteAccessRecord([]byte(wb.Owner))
			if len(rec) != 4+0x70 { // 4 byte header + fixed size payload
				t.Fatalf("WRITEACCESS record is %d bytes, want %d", len(rec), 4+0x70)
			}
			want := n
			if want > maxOwnerLength {
				want = maxOwnerLength
			}
			if got := bytes.TrimRight(rec[4:], " "); len(got) != want {
				t.Errorf("owner stored as %d bytes, want %d", len(got), want)
			}
			if len(wb.GetBiffData()) == 0 {
				t.Error("workbook serialised to nothing")
			}
		})
	}
}

// TestWorkbook_OwnerTruncatesOnRuneBoundary checks that a truncated owner name is
// still valid UTF-8 instead of ending in half a character.
func TestWorkbook_OwnerTruncatesOnRuneBoundary(t *testing.T) {
	// 17 CJK characters are 51 bytes, one character over the 0x30 byte limit.
	wb := NewWorkbook()
	wb.Owner = strings.Repeat("中", 17)

	rec := WriteAccessRecord([]byte(wb.Owner))
	if len(rec) != 4+0x70 {
		t.Fatalf("WRITEACCESS record is %d bytes, want %d", len(rec), 4+0x70)
	}
	name := bytes.TrimRight(rec[4:], " ")
	if !utf8.Valid(name) {
		t.Errorf("truncated owner %q is not valid UTF-8", name)
	}
	if want := 16 * 3; len(name) != want {
		t.Errorf("truncated owner is %d bytes, want %d", len(name), want)
	}
}
