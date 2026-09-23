package xlwt

import (
	"bytes"
	"encoding/binary"
)

func BlankRecord(row int, col int, xfIDX int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(6))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(row))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(col))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(xfIDX))

	return NewBiffRecord(0x0201, buf.Bytes()).Get()
}

func LabelSSTRecord(row int, col int, xfIDX int, sstIDX int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(row))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(col))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(xfIDX))
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(sstIDX))
	return NewBiffRecord(0x00FD, buf.Bytes()).Get()
}

func CalcModeRecord(calcMode int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_h(calcMode))
	return NewBiffRecord(0x000D, buf.Bytes()).Get()
}

func CalcCountRecord(calcCount int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(calcCount))
	return NewBiffRecord(0x000C, buf.Bytes()).Get()
}

func RefModeRecord(refMode int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(refMode))
	return NewBiffRecord(0x00F, buf.Bytes()).Get()
}

func IterationRecord(iterationsOn int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(iterationsOn))
	return NewBiffRecord(0x011, buf.Bytes()).Get()
}

func DeltaRecord(delta float64) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_d(delta))
	return NewBiffRecord(0x010, buf.Bytes()).Get()
}

func SaveRecalcRecord(recalc int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(recalc))
	return NewBiffRecord(0x05F, buf.Bytes()).Get()
}

func GutsRecord(rowGutWidth, colGutHeight, rowVisibleLevels, colVisibleLevels int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(rowGutWidth))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(colGutHeight))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(rowVisibleLevels))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(colVisibleLevels))
	return NewBiffRecord(0x0080, buf.Bytes()).Get()
}

func DefaultRowHeightRecord(options, defHeight int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(options))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(defHeight))
	return NewBiffRecord(0x0225, buf.Bytes()).Get()
}

func WSBoolRecord(options int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(options))
	return NewBiffRecord(0x0081, buf.Bytes()).Get()
}

func DimensionsRecord(firstUsedRow, lastUsedRow, firstUsedCol, lastUsedCol int) []byte {
	if firstUsedRow > lastUsedRow || firstUsedCol > lastUsedCol {
		//	Special case: empty worksheet
		firstUsedRow = 0
		firstUsedCol = 0
		lastUsedRow = -1
		lastUsedCol = -1
	}
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(firstUsedRow))
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(lastUsedRow))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(firstUsedCol))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(lastUsedCol))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	return NewBiffRecord(0x0200, buf.Bytes()).Get()
}

func PrintHeadersRecord(printHeaders int) []byte {
	return SingleHRecord(0x02A, printHeaders)
}

func PrintGridLinesRecord(arg1 int) []byte {
	return SingleHRecord(0x02B, arg1)
}

func GridSetRecord(arg1 int) []byte {
	return SingleHRecord(0x082, arg1)
}

func HorizontalPageBreaksRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	return NewBiffRecord(0x001B, buf.Bytes()).Get()
}

func VerticalPageBreaksRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	return NewBiffRecord(0x001A, buf.Bytes()).Get()
}

func HeaderRecord(headerStr string) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, ASCIIStringPack2(headerStr))
	return NewBiffRecord(0x0014, buf.Bytes()).Get()
}

func FooterRecord(footerStr string) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, ASCIIStringPack2(footerStr))
	return NewBiffRecord(0x0015, buf.Bytes()).Get()
}

func HCenterRecord(arg1 int) []byte {
	return SingleHRecord(0x0083, arg1)
}

func VCenterRecord(arg1 int) []byte {
	return SingleHRecord(0x0084, arg1)
}

func LeftMarginRecord(arg1 float64) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_d(arg1))
	return NewBiffRecord(0x0026, buf.Bytes()).Get()
}

func RightMarginRecord(arg1 float64) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_d(arg1))
	return NewBiffRecord(0x0027, buf.Bytes()).Get()
}

func TopMarginRecord(arg1 float64) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_d(arg1))
	return NewBiffRecord(0x0028, buf.Bytes()).Get()
}

func BottomMarginRecord(arg1 float64) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_d(arg1))
	return NewBiffRecord(0x0029, buf.Bytes()).Get()
}

func SetupPageRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(9))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(100))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(1))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(1))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(1))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x83))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x012C))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x012C))
	_ = binary.Write(&buf, binary.LittleEndian, SP_d(0.1))
	_ = binary.Write(&buf, binary.LittleEndian, SP_d(0.1))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(1))
	return NewBiffRecord(0x00A1, buf.Bytes()).Get()
}

func ScenProtectRecord(arg1 int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(arg1))
	return NewBiffRecord(0x00DD, buf.Bytes()).Get()
}

func RowRecord(index, firstCol, lastCol, heightOptions, options int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(index))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(firstCol))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(lastCol))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(heightOptions))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(options))

	return NewBiffRecord(0x0208, buf.Bytes()).Get()
}

func Window2Record(options, firstVisibleRow, firstVisibleCol, gridColour, previewMagn, normalMagn int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(options))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(firstVisibleRow))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(firstVisibleCol))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(gridColour))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(previewMagn))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(normalMagn))
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))
	// skip scl_magn

	return NewBiffRecord(0x023E, buf.Bytes()).Get()
}

func DefaultWindow2Record() []byte {
	options := 0x02 | 0x04 | 0x10 | 0x20 | 0x80

	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(options))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x40))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))
	// skip scl_magn

	return NewBiffRecord(0x023E, buf.Bytes()).Get()
}

// ColInfoRecord writes a COLINFO record (0x007D) that sets the width, default
// style and outline options of a range of columns.
//
// Layout: first column, last column, width in 1/256 of the width of the zero
// character of the default font, XF index, option flags, unused word.
func ColInfoRecord(firstCol, lastCol, width, xfIndex, options, unused int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(firstCol))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(lastCol))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(width))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(xfIndex))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(options))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(unused))
	return NewBiffRecord(0x007D, buf.Bytes()).Get()
}

// DefColWidthRecord writes a DEFCOLWIDTH record (0x0055), the width used for
// columns that have no COLINFO record.
func DefColWidthRecord(defWidth int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(defWidth))
	return NewBiffRecord(0x0055, buf.Bytes()).Get()
}
