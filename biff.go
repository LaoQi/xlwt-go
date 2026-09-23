package xlwt

import (
	"bytes"
	"encoding/binary"
)

const (
	// stream types
	Biff8BOFRecord__BOOK_GLOBAL = 0x0005
	Biff8BOFRecord__VB_MODULE   = 0x0006
	Biff8BOFRecord__WORKSHEET   = 0x0010
	Biff8BOFRecord__CHART       = 0x0020
	Biff8BOFRecord__MACROSHEET  = 0x0040
	Biff8BOFRecord__WORKSPACE   = 0x0100
)

type BiffRecord struct {
	_REC_ID   SP_H
	_rec_data []byte
}

func NewBiffRecord(RecID SP_H, RecData []byte) *BiffRecord {
	return &BiffRecord{
		_REC_ID:   RecID,
		_rec_data: RecData,
	}
}

func (r *BiffRecord) GetRecHeader() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, r._REC_ID)
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(r._rec_data)))
	return buf.Bytes()
}

// Get serialises the record. A payload longer than 0x2020 bytes is split into a
// chain of CONTINUE (0x003C) records, one payload per record.
//
// The split is a plain byte split and is only valid for records whose data is a
// flat byte string. Records with structured payloads need their own rules: a
// shared string table, for example, must not be cut in the middle of a UTF-16
// character and has to repeat option flags in every continuation, which is why
// SharedStringTable.GetBiffRecord builds its own record chain.
func (r *BiffRecord) Get() []byte {
	data := r._rec_data
	// limit for BIFF7/8: a record payload is at most 0x2020 bytes, everything
	// beyond that goes into CONTINUE records.
	if len(data) > 0x2020 {
		var chunks [][]byte
		pos := 0
		for pos < len(data) {
			// The last chunk is shorter than the limit; taking a fixed-size
			// slice here used to panic for any payload whose length was not a
			// multiple of 0x2020.
			chunkPos := pos + 0x2020
			if chunkPos > len(data) {
				chunkPos = len(data)
			}
			chunks = append(chunks, data[pos:chunkPos])
			pos = chunkPos
		}
		var continuesBuff bytes.Buffer
		_ = binary.Write(&continuesBuff, binary.LittleEndian, r._REC_ID)
		_ = binary.Write(&continuesBuff, binary.LittleEndian, SP_H(len(chunks[0])))
		continuesBuff.Write(chunks[0])

		for _, chunk := range chunks[1:] {
			_ = binary.Write(&continuesBuff, binary.LittleEndian, SP_H(0x003C))
			_ = binary.Write(&continuesBuff, binary.LittleEndian, SP_H(len(chunk)))
			continuesBuff.Write(chunk)
		}
		return continuesBuff.Bytes()
	} else {
		var outBuff bytes.Buffer
		outBuff.Write(r.GetRecHeader())
		outBuff.Write(data)
		return outBuff.Bytes()
	}
}

func SingleHRecord(RecID int, hNum int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(hNum))
	return NewBiffRecord(SP_H(RecID), buf.Bytes()).Get()
}
