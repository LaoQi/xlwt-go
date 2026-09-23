package xlwt

import (
	"bytes"
	"fmt"
	"io"
	"strings"
)

type Workbook struct {
	Owner      string
	Worksheets []*Worksheet
	SST        *SharedStringTable
	Styles     *styleCollection
}

func NewWorkbook() *Workbook {
	return &Workbook{
		Owner:      "None",
		Worksheets: []*Worksheet{},
		SST:        NewSharedStringTable(),
		Styles:     newStyleCollection(),
	}
}

// AddStyle registers a style and returns the XF index that cells using it
// should reference. It is called automatically by Worksheet.Write; call it
// directly only if you need the index itself. A nil style means the default
// style. Registering the same style twice returns the same index.
func (wb *Workbook) AddStyle(style *XFStyle) (int, error) {
	return wb.Styles.AddStyle(style)
}

// AddSheet appends a new worksheet with the given name and returns it.
//
// It returns [ErrInvalidSheetName] for an empty name, a name longer than
// [MaxSheetNameLength] characters, or a name containing \ / ? * [ ] : or a
// leading apostrophe, and [ErrDuplicateSheetName] when another sheet of this
// workbook already uses the name (comparison is case-insensitive).
func (wb *Workbook) AddSheet(name string) (*Worksheet, error) {
	if err := validateSheetName(name); err != nil {
		return nil, err
	}
	lower := strings.ToLower(name)
	for _, existing := range wb.Worksheets {
		if strings.ToLower(existing.Name) == lower {
			return nil, fmt.Errorf("%w: %q", ErrDuplicateSheetName, name)
		}
	}

	ws := NewWorksheet(name, wb.SST)
	ws.workbook = wb
	wb.Worksheets = append(wb.Worksheets, ws)
	return ws, nil
}

// validateSheetName mirrors the worksheet name rules of the file format.
func validateSheetName(name string) error {
	if name == "" {
		return fmt.Errorf("%w: name must not be empty", ErrInvalidSheetName)
	}
	runes := []rune(name)
	if len(runes) > MaxSheetNameLength {
		return fmt.Errorf("%w: %d characters, allowed maximum is %d", ErrInvalidSheetName, len(runes), MaxSheetNameLength)
	}
	if runes[0] == '\'' {
		return fmt.Errorf("%w: name must not start with an apostrophe", ErrInvalidSheetName)
	}
	if strings.ContainsAny(name, "\\/:*?[]") || strings.ContainsRune(name, 0x00) {
		return fmt.Errorf("%w: name must not contain \\ / ? * [ ] : or a NUL byte", ErrInvalidSheetName)
	}
	return nil
}

func (wb *Workbook) BoundsSheetsRec(start int, sheetsLen []int) []byte {
	var prepare bytes.Buffer
	for _, sheet := range wb.Worksheets {
		prepare.Write(BoundSheetRecord(0, 0, sheet.Name))
	}
	start = start + prepare.Len()
	var buf bytes.Buffer
	for i, sheet := range wb.Worksheets {
		buf.Write(BoundSheetRecord(start, 0, sheet.Name))
		if i < len(sheetsLen) {
			// Each worksheet's BOF starts right after the previous one, so the
			// stream position advances by the previous sheet's BIFF length.
			start += sheetsLen[i]
		}
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
	before.Write(wb.Styles.GetBiffData())
	before.Write(PaletteRecord())
	before.Write(UseSelfsRecord())

	// Cells are written sorted and their string references are resolved against
	// the shared string table, so the table must be sorted before any record
	// that quotes an index is built. This is the earliest point at which the
	// whole workbook is known.
	wb.finalizeSST()

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
