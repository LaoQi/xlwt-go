package xlwt

import (
	"bytes"
	"encoding/binary"
)

// Serialisation of the style section of a workbook: FONT, FORMAT, XF and STYLE
// records. The byte layout matches the Python original exactly; a default
// workbook produces the same bytes as before styles were configurable.

// fontRecord writes a FONT record (0x0031).
//
// Layout: height, options, colour index, weight, escapement (5 * uint16),
// underline, family, charset, padding (4 * uint8), name (8-bit length string).
func fontRecord(f Font) []byte {
	f = f.withDefaults()

	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(f.Height))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(f.options()))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(f.ColourIndex))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(f.effectiveWeight()))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(f.Escapement))

	_ = binary.Write(&buf, binary.LittleEndian, SP_B(f.Underline))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(f.Family))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(f.Charset))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00)) // padding

	_ = binary.Write(&buf, binary.LittleEndian, ASCIIStringPack(f.Name))

	return NewBiffRecord(0x0031, buf.Bytes()).Get()
}

// numberFormatRecord writes a FORMAT record (0x041E) for a user defined number
// format string.
func numberFormatRecord(idx int, str string) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(idx))
	_ = binary.Write(&buf, binary.LittleEndian, ASCIIStringPack2(str))
	return NewBiffRecord(0x041E, buf.Bytes()).Get()
}

// xfRecord writes an XF record (0x00E0). When styleXF is true the record
// describes a style XF (used by the built-in styles), otherwise a cell XF.
//
// Layout: font index, format index, protection/type, alignment, rotation, text
// flags, used-attributes, two border words, pattern word.
func xfRecord(fontIdx, numFormatIdx int, a Alignment, b Borders, p Pattern, prot Protection, styleXF bool) []byte {
	var buf bytes.Buffer

	var protBits int
	var usedAttr int
	if styleXF {
		protBits = 0xFFF5
		usedAttr = 0xF4
	} else {
		protBits = (prot.cellLockedFlag() & 0x01) | ((prot.formulaHiddenFlag() & 0x01) << 1)
		usedAttr = 0xF8
	}

	alignByte := ((a.Horz & 0x07) << 0) |
		((a.wrapFlag() & 0x01) << 3) |
		((a.Vert & 0x07) << 4)
	textByte := ((a.Indent & 0x0F) << 0) |
		((a.shrinkFlag() & 0x01) << 4) |
		((a.mergeFlag() & 0x01) << 5) |
		((a.Dire & 0x03) << 6)

	// A line without a style must not carry a colour.
	leftColour, rightColour := b.LeftColour, b.RightColour
	topColour, bottomColour, diagColour := b.TopColour, b.BottomColour, b.DiagColour
	if b.Left == BorderNone {
		leftColour = 0x00
	}
	if b.Right == BorderNone {
		rightColour = 0x00
	}
	if b.Top == BorderNone {
		topColour = 0x00
	}
	if b.Bottom == BorderNone {
		bottomColour = 0x00
	}
	if b.Diag == BorderNone {
		diagColour = 0x00
	}

	diag1, diag2 := 0, 0
	if b.NeedDiag1 {
		diag1 = 1
	}
	if b.NeedDiag2 {
		diag2 = 1
	}

	border1 := ((b.Left & 0x0F) << 0) |
		((b.Right & 0x0F) << 4) |
		((b.Top & 0x0F) << 8) |
		((b.Bottom & 0x0F) << 12) |
		((leftColour & 0x7F) << 16) |
		((rightColour & 0x7F) << 23) |
		((diag1 & 0x01) << 30) |
		((diag2 & 0x01) << 31)

	border2 := ((topColour & 0x7F) << 0) |
		((bottomColour & 0x7F) << 7) |
		((diagColour & 0x7F) << 14) |
		((b.Diag & 0x0F) << 21) |
		((p.Pattern & 0x3F) << 26)

	pattern := ((p.ForeColour & 0x7F) << 0) |
		((p.BackColour & 0x7F) << 7)

	_ = binary.Write(&buf, binary.LittleEndian, SP_H(fontIdx))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(numFormatIdx))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(protBits))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(alignByte))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(a.Rota))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(textByte))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(usedAttr))
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(border1))
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(border2))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(pattern))

	return NewBiffRecord(0x00E0, buf.Bytes()).Get()
}

// styleRecord writes the STYLE record (0x0293) that defines the built-in
// "Normal" cell style.
func styleRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x8000))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0xFF))
	return NewBiffRecord(0x0293, buf.Bytes()).Get()
}
