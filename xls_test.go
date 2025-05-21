package xlwt

import (
	"fmt"
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

	for i := 0; i < 10; i++ {
		for j := 0; j < 10; j++ {
			ws.Write(i, j, fmt.Sprintf("测试输入%d-%d", i, j))
		}
	}

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
