package xlwt_test

import (
	"bytes"
	"fmt"
	"log"
	"os"

	xlwt "github.com/LaoQi/xlwt-go"
)

// ExampleWorkbook shows the minimal workflow: build a workbook, write some
// cells and save it to a file.
func ExampleWorkbook() {
	wb := xlwt.NewWorkbook()
	ws, err := wb.AddSheet("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	if err := ws.Write(0, 0, "Hello, XLS!"); err != nil {
		log.Fatal(err)
	}
	if err := ws.Write(1, 0, "A second row"); err != nil {
		log.Fatal(err)
	}

	fp, err := os.Create("example.xls")
	if err != nil {
		log.Fatal(err)
	}
	defer fp.Close()

	if err := wb.Save(fp); err != nil {
		log.Fatal(err)
	}
}

// ExampleWorkbook_Save writes a workbook into memory and checks that the
// result starts with the OLE2 compound document signature.
func ExampleWorkbook_Save() {
	wb := xlwt.NewWorkbook()
	ws, err := wb.AddSheet("Sheet1")
	if err != nil {
		fmt.Println("add sheet failed:", err)
		return
	}
	if err := ws.Write(0, 0, "Hello, XLS!"); err != nil {
		fmt.Println("write failed:", err)
		return
	}

	var buf bytes.Buffer
	if err := wb.Save(&buf); err != nil {
		fmt.Println("save failed:", err)
		return
	}

	// Every .xls file starts with the OLE2 magic number D0 CF 11 E0 A1 B1 1A E1.
	signature := []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
	fmt.Println("valid xls:", bytes.HasPrefix(buf.Bytes(), signature))
	// Output: valid xls: true
}
