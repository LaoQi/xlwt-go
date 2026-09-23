package xlwt

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"os"
	"path/filepath"
	"testing"
)

// ole2Signature is the magic number every .xls (OLE2 compound document) starts with.
var ole2Signature = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

func TestXlsDoc_Save(t *testing.T) {
	path := filepath.Join(t.TempDir(), "empty.xls")

	fp, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	defer fp.Close()

	if err := NewXlsDoc().Save(fp, []byte{}); err != nil {
		t.Fatal(err)
	}
	if err := fp.Close(); err != nil {
		t.Fatal(err)
	}

	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.HasPrefix(data, ole2Signature) {
		t.Fatalf("output does not start with the OLE2 signature: % X", data[:8])
	}
	if len(data)%SECTOR_SIZE != 0 {
		t.Errorf("output size %d is not a multiple of the sector size %d", len(data), SECTOR_SIZE)
	}
}

func TestWorkbook_Save(t *testing.T) {
	path := filepath.Join(t.TempDir(), "test-go.xls")

	wb := NewWorkbook()
	ws, err := wb.AddSheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}

	for i := 0; i < 10; i++ {
		for j := 0; j < 10; j++ {
			if err := ws.Write(i, j, fmt.Sprintf("测试输入%d-%d", i, j)); err != nil {
				t.Fatal(err)
			}
		}
	}

	fp, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := wb.Save(fp); err != nil {
		fp.Close()
		t.Fatal(err)
	}
	if err := fp.Close(); err != nil {
		t.Fatal(err)
	}

	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.HasPrefix(data, ole2Signature) {
		t.Fatalf("output does not start with the OLE2 signature: % X", data[:8])
	}
	if len(data) < MIN_LIMIT {
		t.Errorf("output size %d is smaller than the OLE2 minimum %d", len(data), MIN_LIMIT)
	}
}

func TestWorksheet_Write(t *testing.T) {
	sst := NewSharedStringTable()
	ws := NewWorksheet("Sheet1", sst)

	for _, w := range []struct {
		r, c int
		s    string
	}{
		{1, 2, "hello"},
		{1, 3, "hello"}, // duplicate string, must reuse the SST index
		{5, 0, "world"},
	} {
		if err := ws.Write(w.r, w.c, w.s); err != nil {
			t.Fatal(err)
		}
	}

	if got, want := len(ws.Grid), 3; got != want {
		t.Errorf("Grid size = %d, want %d", got, want)
	}
	if !ws.RowsIndex[1] || !ws.RowsIndex[5] || len(ws.RowsIndex) != 2 {
		t.Errorf("RowsIndex = %v, want rows 1 and 5", ws.RowsIndex)
	}
	if got, want := len(sst.StrList), 2; got != want {
		t.Errorf("SST size = %d, want %d (duplicates must be deduplicated)", got, want)
	}

	cell := ws.Grid[uint32(1<<16+2)]
	if cell.Row != 1 || cell.Col != 2 || cell.XFIdx != DefaultCellXFStyle {
		t.Errorf("cell = %+v, want row 1 col 2 xf %d", cell, DefaultCellXFStyle)
	}
	// Both copies of "hello" must point at the same shared string.
	if a, b := ws.Grid[uint32(1<<16+2)].SSTIdx, ws.Grid[uint32(1<<16+3)].SSTIdx; a != b {
		t.Errorf("duplicate strings use different SST indexes: %d != %d", a, b)
	}
}

func TestU16StringPack(t *testing.T) {
	// length (2 bytes) + flag (1 byte) + UTF-16LE code units
	got := U16StringPack("A测")
	if len(got) != 3+2*2 {
		t.Fatalf("U16StringPack length = %d, want %d", len(got), 3+2*2)
	}
	if got[0] != 2 || got[1] != 0 || got[2] != 1 {
		t.Errorf("U16StringPack header = % X, want 02 00 01", got[:3])
	}
	if got[3] != 'A' || got[4] != 0 {
		t.Errorf("U16StringPack body = % X, want 41 00 ...", got[3:])
	}
}

// TestXlsDoc_SaveLargeStream covers the secondary MSAT sector path, which is
// only taken when the workbook stream is large enough for more than 109 SAT
// sectors (roughly 6.8 MB). Porting this path used to drop the loop increment
// of i and panicked with "index out of range" for such workbooks.
func TestXlsDoc_SaveLargeStream(t *testing.T) {
	const streamSize = 8 << 20 // 8 MiB: SAT_sect_count > 109, MSAT_sect_count == 1

	buf := bytes.NewBuffer(make([]byte, 0, streamSize+2*SECTOR_SIZE))
	if err := NewXlsDoc().Save(buf, make([]byte, streamSize)); err != nil {
		t.Fatal(err)
	}
	data := buf.Bytes()

	if !bytes.HasPrefix(data, ole2Signature) {
		t.Fatalf("output does not start with the OLE2 signature: % X", data[:8])
	}
	if len(data)%SECTOR_SIZE != 0 {
		t.Errorf("output size %d is not a multiple of the sector size %d", len(data), SECTOR_SIZE)
	}

	// The header is 512 bytes and holds the first 109 SAT sector numbers.
	// MSAT_start_sid follows the min-stream-size and SSAT fields.
	if len(data) < 512 {
		t.Fatalf("output too small: %d bytes", len(data))
	}
	totalSatSectors := binary.LittleEndian.Uint32(data[44:48])
	msatStartSid := int32(binary.LittleEndian.Uint32(data[68:72]))
	totalMsatSectors := binary.LittleEndian.Uint32(data[72:76])

	if totalSatSectors <= 109 {
		t.Errorf("total SAT sectors = %d, want > 109 (secondary MSAT path not exercised)", totalSatSectors)
	}
	if totalMsatSectors == 0 {
		t.Error("total MSAT sectors = 0, want >= 1")
	}
	if msatStartSid < 0 {
		t.Errorf("MSAT start SID = %d, want a valid sector index", msatStartSid)
	}
	if int(msatStartSid)*SECTOR_SIZE >= len(data) {
		t.Errorf("MSAT start SID %d points past the end of the file (%d bytes)", msatStartSid, len(data))
	}
}
