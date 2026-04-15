package xlwt

import (
	"bytes"
	"sort"
)

const DefaultRowHeightOptions = 0x00FF & 0x07FFF

type Cell struct {
	Row    int
	Col    int
	SSTIdx int
	XFIdx  int
}

type Worksheet struct {
	Name      string
	SST       *SharedStringTable
	Grid      map[uint32]Cell
	RowsIndex map[int]bool
}

func NewWorksheet(name string, sst *SharedStringTable) *Worksheet {
	return &Worksheet{
		Name:      name,
		SST:       sst,
		Grid:      make(map[uint32]Cell),
		RowsIndex: make(map[int]bool),
	}
}

func (ws *Worksheet) calcSettingsRec() []byte {
	var buf bytes.Buffer
	buf.Write(CalcModeRecord(1))
	buf.Write(CalcCountRecord(0x0064))
	buf.Write(RefModeRecord(1))
	buf.Write(IterationRecord(0))
	buf.Write(DeltaRecord(0.001))
	buf.Write(SaveRecalcRecord(0))

	return buf.Bytes()
}

func (ws *Worksheet) GutsRec() []byte {
	//@todo __update_row_visible_levels

	return GutsRecord(0, 0, 1, 0)
}

func (ws *Worksheet) defaultRowHeightRec() []byte {
	return DefaultRowHeightRecord(0x0000, 0x00FF)
}

func (ws *Worksheet) wsBoolRec() []byte {
	options := 0x01
	options |= 0x01 << 10 // __show_row_outline
	options |= 0x01 << 11 // __show_col_outline
	return WSBoolRecord(options)
}

func (ws *Worksheet) colInfoRec() []byte {
	return []byte{}
}

func (ws *Worksheet) dimensionsRec() []byte {
	lastUsedRow := 0
	lastUsedCol := 0
	for _, cell := range ws.Grid {
		if cell.Row > lastUsedRow {
			lastUsedRow = cell.Row
		}
		if cell.Col > lastUsedCol {
			lastUsedCol = cell.Col
		}
	}
	return DimensionsRecord(0, lastUsedRow+1, 0, lastUsedCol+1)
}

func (ws *Worksheet) printSettingsRec() []byte {
	var buf bytes.Buffer
	buf.Write(PrintHeadersRecord(0))
	buf.Write(PrintGridLinesRecord(0))
	buf.Write(GridSetRecord(1))
	buf.Write(HorizontalPageBreaksRecord())
	buf.Write(VerticalPageBreaksRecord())
	buf.Write(HeaderRecord("&P"))
	buf.Write(FooterRecord("&F"))
	buf.Write(HCenterRecord(1))
	buf.Write(VCenterRecord(0))
	buf.Write(LeftMarginRecord(0.3))
	buf.Write(RightMarginRecord(0.3))
	buf.Write(TopMarginRecord(0.61))
	buf.Write(BottomMarginRecord(0.37))
	buf.Write(SetupPageRecord())

	return buf.Bytes()
}

func (ws *Worksheet) protectionRec() []byte {
	var buf bytes.Buffer
	buf.Write(ProtectRecord(0))
	buf.Write(ScenProtectRecord(0))
	buf.Write(WindowProtectRecord(0))
	buf.Write(ObjectProtectRecord(0))
	buf.Write(PasswordRecord(""))
	return buf.Bytes()
}

func (ws *Worksheet) GetRowCellsBiffData(row int) []byte {
	var cells []Cell
	for _, cell := range ws.Grid {
		if cell.Row == row {
			cells = append(cells, cell)
		}
	}
	if len(cells) == 0 {
		return []byte{}
	}
	sort.Slice(cells, func(i, j int) bool {
		return cells[i].Col < cells[j].Col
	})
	firstCol := cells[0].Col
	lastCol := cells[len(cells)-1].Col + 1

	options := (0x01 & 0x01) << 8
	options |= (0x0F & 0x0FFF) << 16 // default style

	var buf bytes.Buffer
	buf.Write(RowRecord(row, firstCol, lastCol, DefaultRowHeightOptions, options))

	for _, cell := range cells {
		buf.Write(LabelSSTRecord(row, cell.Col, cell.XFIdx, cell.SSTIdx))
	}

	return buf.Bytes()
}

func (ws *Worksheet) GetRowsBiffData() []byte {
	var buf bytes.Buffer

	var rows []int
	for index, _ := range ws.RowsIndex {
		rows = append(rows, index)
	}
	sort.Ints(rows)
	for _, index := range rows {
		buf.Write(ws.GetRowCellsBiffData(index))
	}
	return buf.Bytes()
}

func (ws *Worksheet) GetBiffData() []byte {
	var buf bytes.Buffer
	buf.Write(Biff8BOFRecord(Biff8BOFRecord__WORKSHEET))
	buf.Write(ws.calcSettingsRec())
	buf.Write(ws.GutsRec())
	buf.Write(ws.defaultRowHeightRec())
	buf.Write(ws.wsBoolRec())
	buf.Write(ws.colInfoRec())
	buf.Write(ws.dimensionsRec())
	buf.Write(ws.printSettingsRec())
	buf.Write(ws.protectionRec())

	buf.Write(ws.GetRowsBiffData())

	// skip MergedCellsRecord ObjBmpRecord
	buf.Write(DefaultWindow2Record())
	// skip PanesRecord
	buf.Write(EOFRecord())

	return buf.Bytes()
}

func (ws *Worksheet) Write(r, c int, label string) {
	var key uint32
	key = uint32((r << 16) + c)
	idx := ws.SST.AddStr(label)
	ws.Grid[key] = Cell{Row: r, Col: c, SSTIdx: idx, XFIdx: DefaultCellXFStyle}
	ws.RowsIndex[r] = true
}
