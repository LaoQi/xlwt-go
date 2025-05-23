package xlwt

import (
	"bytes"
	"encoding/binary"
	"log"
)

const MaxSSTLength = 0x2020
const MaxSSTCellLength = 0x2000
const CONTINUE_ID = 0x003C

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

func (sst *SharedStringTable) GetBiffRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x00FC)) // SST
	//firstLen, body := sst.GetSSTRecord()

	var tmp bytes.Buffer
	firstLen := 8
	hasWriteFirst := false
	for _, str := range sst.StrList {
		//buf.Write(U16StringPack(str))
		block := U16StringPack(str)
		if len(block) >= MaxSSTCellLength {
			log.Println("exceed max sst cell length", len(block))
			continue
		}
		if tmp.Len()+len(block) > MaxSSTLength {
			if !hasWriteFirst {
				hasWriteFirst = true
				firstLen = firstLen + tmp.Len()
				_ = binary.Write(&buf, binary.LittleEndian, SP_H(firstLen))
				_ = binary.Write(&buf, binary.LittleEndian, SP_I(sst.Total))
				_ = binary.Write(&buf, binary.LittleEndian, SP_I(len(sst.StrList)))
			} else {
				//	write continue
				_ = binary.Write(&buf, binary.LittleEndian, SP_H(CONTINUE_ID))
				_ = binary.Write(&buf, binary.LittleEndian, SP_H(tmp.Len()))
			}
			buf.Write(tmp.Bytes())
			tmp.Reset()
		}
		tmp.Write(block)
	}

	if !hasWriteFirst {
		firstLen = firstLen + tmp.Len()
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(firstLen))
		_ = binary.Write(&buf, binary.LittleEndian, SP_I(sst.Total))
		_ = binary.Write(&buf, binary.LittleEndian, SP_I(len(sst.StrList)))
	} else {
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(CONTINUE_ID))
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(tmp.Len()))
	}
	buf.Write(tmp.Bytes())

	return buf.Bytes()
}
