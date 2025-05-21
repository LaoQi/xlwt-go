package xlwt

import (
	"bytes"
	"encoding/binary"
)

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

func (sst *SharedStringTable) GetSSTRecord() []byte {
	var buf bytes.Buffer
	for _, str := range sst.StrList {
		buf.Write(U16StringPack(str))
	}
	return buf.Bytes()
}

func (sst *SharedStringTable) GetBiffRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x00FC)) // SST
	body := sst.GetSSTRecord()
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(body)+8))
	_ = binary.Write(&buf, binary.LittleEndian, SP_I(sst.Total))
	_ = binary.Write(&buf, binary.LittleEndian, SP_I(len(sst.StrList)))
	buf.Write(body)
	return buf.Bytes()
}
