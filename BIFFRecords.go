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

func (self *BiffRecord) GetRecHeader() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, self._REC_ID)
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(self._rec_data)))
	return buf.Bytes()
}

func (self *BiffRecord) Get() []byte {
	data := self._rec_data
	// limit for BIFF7/8
	if len(data) > 0x2020 {
		var chunks [][]byte
		pos := 0
		for pos < len(data) {
			chunk_pos := pos + 0x2020
			chunk := data[pos:chunk_pos]
			chunks = append(chunks, chunk)
			pos = chunk_pos
		}
		var continuesBuff bytes.Buffer
		_ = binary.Write(&continuesBuff, binary.LittleEndian, self._REC_ID)
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
		outBuff.Write(self.GetRecHeader())
		outBuff.Write(data)
		return outBuff.Bytes()
	}
}

func SingleHRecord(RecID int, hNum int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(hNum))
	return NewBiffRecord(SP_H(RecID), buf.Bytes()).Get()
}

func Biff8BOFRecord(recType SP_H) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x0600)) // version
	_ = binary.Write(&buf, binary.LittleEndian, recType)
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x0DBB)) // build
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x07CC)) // year
	_ = binary.Write(&buf, binary.LittleEndian, SP_I(0x00))   // file_hist_flags
	_ = binary.Write(&buf, binary.LittleEndian, SP_I(0x06))   // ver_can_read

	return NewBiffRecord(0x0809, buf.Bytes()).Get()
}

func InteraceHdrRecord() []byte {
	return NewBiffRecord(0x00E1, []byte{0xB0, 0x04}).Get()
}

func InteraceEndRecord() []byte {
	return NewBiffRecord(0x00E2, []byte{}).Get()
}

func MMSRecord() []byte {
	return NewBiffRecord(0x00C1, SP_H_0).Get()
}

func CodepageBiff8Record() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0x04B0)) // UTF16
	return NewBiffRecord(0x0042, buf.Bytes()).Get()
}

func DSFRecord() []byte {
	return NewBiffRecord(0x0161, SP_H_0).Get()
}

func TabIDRecord(sheetCount int) []byte {
	var buf bytes.Buffer
	for i := 0; i < sheetCount; i++ {
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(i+1))
	}
	return NewBiffRecord(0x013D, buf.Bytes()).Get()
}

func FnGroupCountRecord() []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x0E))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00))
	return NewBiffRecord(0x009C, buf.Bytes()).Get()
}

func WindowProtectRecord(wndprotect int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(wndprotect))
	return NewBiffRecord(0x0019, buf.Bytes()).Get()
}

func ProtectRecord(protect int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(protect))
	return NewBiffRecord(0x0012, buf.Bytes()).Get()
}

func ObjectProtectRecord(objprotect int) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(objprotect))
	return NewBiffRecord(0x0063, buf.Bytes()).Get()
}

func PasswordRecord(_ string) []byte {
	//@fixme not support passwd
	return NewBiffRecord(0x0013, SP_H_0).Get()
}

func Prot4RevRecord() []byte {
	return NewBiffRecord(0x01AF, SP_H_0).Get()
}

func Prot4RevPassRecord() []byte {
	return NewBiffRecord(0x01BC, SP_H_0).Get()
}

// BackupRecord backup int
func BackupRecord(_ int) []byte {
	//@fixme not support backup
	//This  record  contains  a Boolean value determining whether Excel makes
	//    a backup of the file while saving.
	return NewBiffRecord(0x0040, SP_H_0).Get()
}

func HideObjRecord() []byte {
	return NewBiffRecord(0x008D, SP_H_0).Get()
}

func Window1Record() []byte {
	//@fixme not support Window1Record
	// pack('<9H')
	return NewBiffRecord(0x003D, []byte{
		0xE0, 0x01, 0x5A, 0x00, 0xCF, 0x3F, 0x4E, 0x2A, 0x38,
		0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x58, 0x02,
	}).Get()
}

func DateModeRecord(from1904 bool) []byte {
	var buf bytes.Buffer
	if from1904 {
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(1))
	} else {
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	}
	return NewBiffRecord(0x0022, buf.Bytes()).Get()
}

func PrecisionRecord(use_real_values bool) []byte {
	var buf bytes.Buffer
	if use_real_values {
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(1))
	} else {
		_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	}
	return NewBiffRecord(0x000E, buf.Bytes()).Get()
}

func RefreshAllRecord() []byte {
	return NewBiffRecord(0x01B7, SP_H_0).Get()
}

func BookBoolRecord() []byte {
	return NewBiffRecord(0x00DA, SP_H_0).Get()
}

func PaletteRecord() []byte {
	//@fixme not support
	//return NewBiffRecord(0x0092, []byte{0x00}).Get()
	return []byte{}
}

//func CountryRecord() []byte {
//	//@fixme not support
//	//return NewBiffRecord(0x008C, []byte{0x00}).Get()
//	return []byte{}
//}

func UseSelfsRecord() []byte {
	return NewBiffRecord(0x0160, SP_H_1).Get()
}

func BoundSheetRecord(streamPos int, visibility int, sheet string) []byte {
	var buf bytes.Buffer
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(streamPos))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(visibility))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0))
	_ = binary.Write(&buf, binary.LittleEndian, ASCIIStringPack(sheet))
	return NewBiffRecord(0x0085, buf.Bytes()).Get()
}

func EOFRecord() []byte {
	return NewBiffRecord(0x000A, []byte{}).Get()
}

func WriteAccessRecord(owner []byte) []byte {
	var buf bytes.Buffer
	paddingLength := 0x70 - len(owner)
	_ = binary.Write(&buf, binary.LittleEndian, owner)
	padding := FillBytes(paddingLength, 0x20)
	_ = binary.Write(&buf, binary.LittleEndian, padding)
	return NewBiffRecord(0x005C, buf.Bytes()).Get()
}

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
	//_ = binary.Write(&buf, binary.LittleEndian, SP_H(6))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(row))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(col))
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(xfIDX))
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(sstIDX))
	return NewBiffRecord(0x00FD, buf.Bytes()).Get()
}

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

	_ = binary.Write(&buf, binary.LittleEndian, ASCIIStringPack("Arial")) // escapement

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

	//_ = binary.Write(&buf, binary.LittleEndian, SP_L(0x20400000)) // borders
	//_ = binary.Write(&buf, binary.LittleEndian, SP_L(0x102040))   // borders
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

	//_ = binary.Write(&buf, binary.LittleEndian, SP_L(0x20400000)) // borders
	//_ = binary.Write(&buf, binary.LittleEndian, SP_L(0x102040))   // borders
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
	//options |= (self.__show_formulas        & 0x01) << 0
	//options |= (self.__show_grid            & 0x01) << 1
	//options |= (self.__show_headers         & 0x01) << 2
	//options |= (self.__panes_frozen         & 0x01) << 3
	//options |= (self.show_zero_values       & 0x01) << 4
	//options |= (self.__auto_colour_grid     & 0x01) << 5
	//options |= (self.__cols_right_to_left   & 0x01) << 6
	//options |= (self.__show_outline         & 0x01) << 7
	//options |= (self.__remove_splits        & 0x01) << 8
	//options |= (self.__selected             & 0x01) << 9
	//options |= (self.__sheet_visible        & 0x01) << 10
	//options |= (self.__page_preview         & 0x01) << 11
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
