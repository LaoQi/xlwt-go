package xlwt

import (
	"bytes"
	"io"
)

type Workbook struct {
	Owner      string
	Worksheets []*Worksheet
}

func NewWorkbook() *Workbook {
	return &Workbook{
		Owner:      "None",
		Worksheets: []*Worksheet{},
	}
}

func (wb *Workbook) AddSheet(name string) *Worksheet {
	ws := NewWorksheet(name)
	wb.Worksheets = append(wb.Worksheets, ws)
	return ws
}

func (wb *Workbook) GetBiffData() []byte {
	var buf bytes.Buffer
	buf.Write(Biff8BOFRecord(Biff8BOFRecord__BOOK_GLOBAL))
	buf.Write(InteraceHdrRecord())
	buf.Write(MMSRecord())
	buf.Write(InteraceEndRecord())

	buf.Write(WriteAccessRecord([]byte(wb.Owner))) // owner
	buf.Write(CodepageBiff8Record())
	buf.Write(DSFRecord())
	buf.Write(TabIDRecord(len(wb.Worksheets)))
	buf.Write(FnGroupCountRecord())
	buf.Write(WindowProtectRecord(0))
	buf.Write(ProtectRecord(0))
	buf.Write(ObjectProtectRecord(0))
	buf.Write(PasswordRecord("0"))
	buf.Write(Prot4RevRecord())
	buf.Write(Prot4RevPassRecord())
	buf.Write(BackupRecord(0))
	buf.Write(HideObjRecord())
	buf.Write(Window1Record())
	buf.Write(DateModeRecord(false))
	buf.Write(PrecisionRecord(true))
	buf.Write(RefreshAllRecord())
	buf.Write(BookBoolRecord())
	//buf.Write(Style())
	buf.Write(PaletteRecord())
	buf.Write(UseSelfsRecord())

	return buf.Bytes()
}

func (wb *Workbook) Save(writer io.Writer) error {
	doc := NewXlsDoc()
	return doc.Save(writer, wb.GetBiffData())
}
