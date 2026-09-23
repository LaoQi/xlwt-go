package xlwt

import (
	"bytes"
	"encoding/binary"
	"strconv"
	"testing"
)

// TestBiffRecord_ChunkSplit covers payload lengths that used to panic: every
// length above 0x2020 that is not a multiple of it took a full-size slice past
// the end of the data.
func TestBiffRecord_ChunkSplit(t *testing.T) {
	for _, n := range []int{
		0, 1, 0x201F, 0x2020, 0x2021, 0x202B,
		0x4041, 0x6061, 0x8000, 2*0x2020 - 1, 2 * 0x2020,
	} {
		name := "len_" + strconv.Itoa(n)
		t.Run(name, func(t *testing.T) {
			payload := make([]byte, n)
			for i := range payload {
				payload[i] = byte(i % 251)
			}
			got := NewBiffRecord(0x00FC, payload).Get()

			// Re-parse the result: the record plus its CONTINUE chain must hold
			// exactly the input bytes.
			var recovered []byte
			pos := 0
			first := true
			for pos < len(got) {
				if pos+4 > len(got) {
					t.Fatalf("truncated record header at %d", pos)
				}
				id := binary.LittleEndian.Uint16(got[pos:])
				size := int(binary.LittleEndian.Uint16(got[pos+2:]))
				if size > 0x2020 {
					t.Fatalf("record at %d claims %d bytes, limit is 0x2020", pos, size)
				}
				wantID := uint16(0x003C)
				if first {
					wantID = 0x00FC
					first = false
				} else if id != wantID {
					t.Fatalf("record at %d has id 0x%04X, want CONTINUE", pos, id)
				}
				if pos+4+size > len(got) {
					t.Fatalf("record at %d overruns the output", pos)
				}
				recovered = append(recovered, got[pos+4:pos+4+size]...)
				pos += 4 + size
			}
			if !bytes.Equal(recovered, payload) {
				t.Errorf("recovered %d bytes, want %d", len(recovered), len(payload))
			}
			if len(payload) <= 0x2020 {
				if want := 4 + len(payload); len(got) != want {
					t.Errorf("single record is %d bytes, want %d", len(got), want)
				}
			}
		})
	}
}
