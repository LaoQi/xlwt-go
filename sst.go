package xlwt

import (
	"bytes"
	"encoding/binary"
)

// MaxSSTLength is the maximum payload size of a single SST or CONTINUE record
// in BIFF8 (0x2020 = 8224 bytes).
const MaxSSTLength = 0x2020

// MaxSSTCellLength is a legacy constant from an earlier porting attempt and is
// no longer used: strings longer than a single record are split across CONTINUE
// records instead of being skipped.
const MaxSSTCellLength = 0x2000

const CONTINUE_ID = 0x003C

const SST_ID = 0x00FC

type SharedStringTable struct {
	StrIndexes map[string]int
	StrList    []string
	Total      int
}

func NewSharedStringTable() *SharedStringTable {
	return &SharedStringTable{
		StrIndexes: make(map[string]int),
	}
}

func (sst *SharedStringTable) AddStr(value string) int {
	// ignore encoding
	sst.Total += 1
	if idx, ok := sst.StrIndexes[value]; ok {
		return idx
	}
	idx := len(sst.StrList)
	sst.StrList = append(sst.StrList, value)
	sst.StrIndexes[value] = idx
	return idx
}

// GetBiffRecord builds the SST record with all shared strings.
//
// The payload of a BIFF8 record is limited to 0x2020 bytes, so it is written as
// a chain of an SST record followed by CONTINUE records. A string that does not
// fit into the remaining space is split across the boundary: the character data
// continues in the next record, and that record starts with a repeated option
// flags byte (as required by BIFF8 for a split Unicode string).
func (sst *SharedStringTable) GetBiffRecord() []byte {
	var (
		records [][]byte // finished record payloads, without the 4-byte record header
		current []byte   // payload of the record currently being built
	)

	// The SST record starts with an 8-byte header holding the total number of
	// string references and the number of unique strings. Both are patched in
	// after all strings have been written, because they must include the bytes
	// in this space when the record is filled up.
	current = make([]byte, 8, MaxSSTLength)

	newPiece := func() {
		records = append(records, current)
		current = nil
	}

	saveAtom := func(atom []byte) {
		if MaxSSTLength-len(current) < len(atom) {
			newPiece()
		}
		current = append(current, atom...)
	}

	// saveChars appends the character data of a string, splitting it across
	// records when it does not fit. unicode selects the option flags byte that
	// has to be repeated at the start of every continuation.
	saveChars := func(chars []byte, unicode bool) {
		flag := byte(0x00)
		if unicode {
			flag = 0x01
		}
		for i := 0; i < len(chars); {
			free := MaxSSTLength - len(current)
			remaining := len(chars) - i
			var n int
			switch {
			case free >= remaining:
				n = remaining
			case unicode:
				n = free & 0xFFFE // keep UTF-16 code units intact
			default:
				n = free
			}
			current = append(current, chars[i:i+n]...)
			i += n
			if i < len(chars) {
				newPiece()
				current = append(current, flag)
			}
		}
	}

	for _, str := range sst.StrList {
		block := U16StringPack(str) // 2-byte length + option flags + UTF-16 code units

		// The string header and the first code unit are kept together so that
		// they never straddle a record boundary.
		atomLen := 5
		if len(block) < atomLen {
			atomLen = len(block) // empty string: nothing but length and flags
		}
		saveAtom(block[:atomLen])
		saveChars(block[atomLen:], block[2] == 0x01)
	}
	newPiece()

	var buf bytes.Buffer
	for i, payload := range records {
		recID := SP_H(CONTINUE_ID)
		if i == 0 {
			recID = SP_H(SST_ID)
		}
		_ = binary.Write(&buf, binary.LittleEndian, recID)
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(payload)))
		buf.Write(payload)
	}

	// Patch the counts into the placeholder at the start of the first record:
	// 2 bytes record id + 2 bytes length, then total and unique counts.
	out := buf.Bytes()
	binary.LittleEndian.PutUint32(out[4:], uint32(sst.Total))
	binary.LittleEndian.PutUint32(out[8:], uint32(len(sst.StrList)))

	return out
}
