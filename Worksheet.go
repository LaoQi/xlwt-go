package xlwt

import "bytes"

type Cell struct {
	Row   int
	Col   int
	Label string
}

type Worksheet struct {
	Name string
	Grid map[uint32]Cell
}

func NewWorksheet(name string) *Worksheet {
	return &Worksheet{
		Name: name,
		Grid: make(map[uint32]Cell),
	}
}

func (ws *Worksheet) GetBiffData() []byte {
	var buf bytes.Buffer
	return buf.Bytes()
}

func (ws *Worksheet) Write(r, c int, label string) {
	var key uint32
	key = uint32((r << 16) + c)
	ws.Grid[key] = Cell{Row: r, Col: c, Label: label}
}
