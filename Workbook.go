package xlwt

import (
	"bytes"
	"io"
)

type Workbook struct {
	Owner      string
	Worksheets []*Worksheet
	SST        *SharedStringTable
	Style      *XFStyle
}

func NewWorkbook() *Workbook {
	return &Workbook{
		Owner:      "None",
		Worksheets: []*Worksheet{},
		SST:        NewSharedStringTable(),
	}
}

func (wb *Workbook) AddStyle(s *XFStyle) {

}

func (wb *Workbook) AddSheet(name string) *Worksheet {
	ws := NewWorksheet(name, wb.SST)
	wb.Worksheets = append(wb.Worksheets, ws)
	return ws
}

func (wb *Workbook) BoundsSheetsRec(start int, sheetsLen []int) []byte {
	var prepare bytes.Buffer
	for _, sheet := range wb.Worksheets {
		prepare.Write(BoundSheetRecord(0, 0, sheet.Name))
	}
	start = start + prepare.Len()
	var buf bytes.Buffer
	for _, sheet := range wb.Worksheets {
		buf.Write(BoundSheetRecord(start, 0, sheet.Name))
	}
	return buf.Bytes()
}

func (wb *Workbook) GetBiffData() []byte {
	var before bytes.Buffer
	before.Write(Biff8BOFRecord(Biff8BOFRecord__BOOK_GLOBAL))
	before.Write(InteraceHdrRecord())
	before.Write(MMSRecord())
	before.Write(InteraceEndRecord())

	before.Write(WriteAccessRecord([]byte(wb.Owner))) // owner
	before.Write(CodepageBiff8Record())
	before.Write(DSFRecord())
	before.Write(TabIDRecord(len(wb.Worksheets)))
	before.Write(FnGroupCountRecord())
	before.Write(WindowProtectRecord(0))
	before.Write(ProtectRecord(0))
	before.Write(ObjectProtectRecord(0))
	before.Write(PasswordRecord("0"))
	before.Write(Prot4RevRecord())
	before.Write(Prot4RevPassRecord())
	before.Write(BackupRecord(0))
	before.Write(HideObjRecord())
	before.Write(Window1Record())
	before.Write(DateModeRecord(false))
	before.Write(PrecisionRecord(true))
	before.Write(RefreshAllRecord())
	before.Write(BookBoolRecord())
	before.Write(wb.Style.GetBiffData())
	before.Write(PaletteRecord())
	before.Write(UseSelfsRecord())

	var after bytes.Buffer
	//after.Write(CountryRec())  // Skip
	//after.Write(LinksRec())  // Skip
	after.Write(wb.SST.GetBiffRecord())

	eof := EOFRecord()

	var sheets bytes.Buffer
	var sheetLength []int
	for _, ws := range wb.Worksheets {
		bd := ws.GetBiffData()
		sheetLength = append(sheetLength, len(bd))
		sheets.Write(bd)
	}

	var out bytes.Buffer

	out.Write(before.Bytes())
	out.Write(wb.BoundsSheetsRec(before.Len()+after.Len()+len(eof), sheetLength))
	out.Write(after.Bytes())
	out.Write(eof)
	out.Write(sheets.Bytes())

	return out.Bytes()
}

func (wb *Workbook) Save(writer io.Writer) error {
	doc := NewXlsDoc()
	return doc.Save(writer, wb.GetBiffData())
}
