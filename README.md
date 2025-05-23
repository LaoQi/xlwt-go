# xlwt-go
Simple Go Implementation of Python xlwt Library

### Notice
* Only support string cell
* Not support style
* Not support formula
* May no longer be maintained

### Use

```bash
$ git clone https://github.com/LaoQi/xlwt-go.git
```

```
// go.mod
replace xlwt => github.com/LaoQi/xlwt-go master 
// or
replace xlwt => ./xlwt-go
```

```go
wb := xlwt.NewWorkbook()
ws := wb.AddSheet("Sheet1")

ws.Write(0, 0, "xls is bullshit!!!")

fp, err := os.Create("test-go.xls")
if err != nil {
    log.Fatal(err)
}
defer fp.Close()
err = wb.Save(fp)
```