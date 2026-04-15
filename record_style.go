package xlwt

import (
	"bytes"
	"encoding/binary"
)

func DefaultFontRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x00C8)) // height
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x00))   // options
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x7FFF)) // colour_index
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x0190)) // weight
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x00))   // escapement

	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00)) // underline
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00)) // family
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x01)) // charset
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00)) // padding

	_ = binary.Write(&buf, binary.LittleEndian, ASCIIStringPack("Arial"))

	return NewBiffRecord(0x0031, buf.Bytes()).Get()
}

func NumberFormatRecord(idx int, str string) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(idx))
	_ = binary.Write(&buf, binary.LittleEndian, ASCIIStringPack2(str))
	return NewBiffRecord(0x041E, buf.Bytes()).Get()
}

func CellXFRecord(fontIdx int) []byte {
	var buf bytes.Buffer

	_ = binary.Write(&buf, binary.LittleEndian, SP_H(fontIdx))                           // font_xf_idx
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(FIRST_USER_DEFINED_NUM_FORMAT_IDX)) // fmt_str_xf_idx
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(1))                                 // cell
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(2<<4))                              // alignment
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0))                                 // ROTATION
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0))                                 // txt
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0xF8))                              // cell

	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0)) // borders
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0)) // borders

	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x20C0)) // pattern

	return NewBiffRecord(0x00E0, buf.Bytes()).Get()
}

func DefaultCellXFRecord() []byte {
	return CellXFRecord(6)
}

func DefaultXFRecord() []byte {
	var buf bytes.Buffer

	_ = binary.Write(&buf, binary.LittleEndian, SP_H(6))                                 // font_xf_idx
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(FIRST_USER_DEFINED_NUM_FORMAT_IDX)) // fmt_str_xf_idx
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0xFFF5))                            // not cell
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(2<<4))                              // alignment
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0))                                 // ROTATION
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0))                                 // txt
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0xF4))                              // not cell

	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0)) // borders
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0)) // borders

	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x20C0)) // pattern

	return NewBiffRecord(0x00E0, buf.Bytes()).Get()
}

func StyleRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x8000))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0xFF))
	return NewBiffRecord(0x0293, buf.Bytes()).Get()
}
