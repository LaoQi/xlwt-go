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

// U16StringPack1Byte packs a string with a 1-byte character count, as used by
// record types that cannot hold more than 255 characters (for example the sheet
// name of a BOUNDSHEET record).
//
// Strings that fit into latin-1 use the compressed 8-bit format, everything
// else is written as UTF-16LE with the high-byte flag set. This mirrors the
// upack1 helper of the Python original.
func U16StringPack1Byte(s string) []byte {
	runes := []rune(s)

	compressed := true
	for _, r := range runes {
		if r > 0xFF {
			compressed = false
			break
		}
	}

	var buf bytes.Buffer
	if compressed {
		_ = binary.Write(&buf, binary.LittleEndian, SP_B(len(runes)))
		_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00))
		for _, r := range runes {
			_ = binary.Write(&buf, binary.LittleEndian, SP_B(byte(r)))
		}
		return buf.Bytes()
	}

	u16s := utf16.Encode(runes)
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(len(u16s)))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x01))
	for _, u := range u16s {
		_ = binary.Write(&buf, binary.LittleEndian, u)
	}
	return buf.Bytes()
}
