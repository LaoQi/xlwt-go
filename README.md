# xlwt-go

A Go library for writing Microsoft Excel (.xls) files. Ported from the Python [xlwt](https://github.com/python-excel/xlwt) library.

## Install

```bash
go get github.com/LaoQi/xlwt-go
```

## Usage

```go
package main

import (
	"log"
	"os"

	xlwt "github.com/LaoQi/xlwt-go"
)

func main() {
	wb := xlwt.NewWorkbook()
	ws := wb.AddSheet("Sheet1")

	ws.Write(0, 0, "Hello, XLS!")

	fp, err := os.Create("output.xls")
	if err != nil {
		log.Fatal(err)
	}
	defer fp.Close()

	if err := wb.Save(fp); err != nil {
		log.Fatal(err)
	}
}
```

## Limitations

- Only supports string cells
- No style customization
- No formula support
- No number/date cell types

## License

LGPL-2.1
