package xlwt

import (
	"bytes"
	"encoding/binary"
	"unicode/utf16"
)

func FillBytes(size int, value byte) []byte {
	buf := make([]byte, size)
	for i := 0; i < size; i++ {
		buf[i] = value
	}
	return buf
}

func FillInt(size int, value int) []int {
	buf := make([]int, size)
	for i := 0; i < size; i++ {
		buf[i] = value
	}
	return buf
}

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

func ASCIIStringPack(s string) []byte {
	var buf bytes.Buffer
	sb := []byte(s)
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(len(sb)))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x0)) // flag
	_ = binary.Write(&buf, binary.LittleEndian, sb)
	return buf.Bytes()
}

func ASCIIStringPack2(s string) []byte {
	var buf bytes.Buffer
	sb := []byte(s)
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(sb)))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x0)) // flag
	_ = binary.Write(&buf, binary.LittleEndian, sb)
	return buf.Bytes()
}
