package xlwt

import (
	"log"
	"os"
	"testing"
)

func TestXlsDoc_Save(t *testing.T) {
	xlsDoc := NewXlsDoc()
	fp, err := os.Create("test-go.xls")
	if err != nil {
		log.Fatal(err)
	}
	defer fp.Close()
	err = xlsDoc.Save(fp, []byte{})
	if err != nil {
		log.Fatal(err)
	}
}

func TestWorkbook_Save(t *testing.T) {
	wb := NewWorkbook()
	ws := wb.AddSheet("Sheet1")
	ws.Write(0, 0, "测试测试")
	fp, err := os.Create("test-go.xls")
	if err != nil {
		log.Fatal(err)
	}
	defer fp.Close()
	err = wb.Save(fp)
	if err != nil {
		log.Fatal(err)
	}
}
