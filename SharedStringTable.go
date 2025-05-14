package xlwt

import (
	"bytes"
	"encoding/binary"
	"unicode/utf16"
)

func U16StringPack(s string) []byte {
	var buf bytes.Buffer
	// always use Unicode 16
	u16s := utf16.Encode([]rune(s))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(u16s)))
	_ = binary.Write(&buf, binary.LittleEndian, uint8(1))
	for _, u := range u16s {
		_ = binary.Write(&buf, binary.LittleEndian, u)
	}
	return buf.Bytes()
}

type SharedStringTable struct {
	StrIndexes map[string]int
	StrList    []string
	Total      int
}

func NewSharedStringTable() *SharedStringTable {
	return &SharedStringTable{}
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
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x00fc)) // SST
	body := sst.GetSSTRecord()
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(body)+8))
	_ = binary.Write(&buf, binary.LittleEndian, SP_I(sst.Total))
	_ = binary.Write(&buf, binary.LittleEndian, SP_I(len(sst.StrList)))
	buf.Write(body)
	return buf.Bytes()
}
