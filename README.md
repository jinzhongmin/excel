# excel 基于go-ole


可以操作xls文件,必须先安装Microsoft office excel,2010以前的版本最多只能操作65535行。

为了便于操作，任何函数不会因抛出错误而退出程序，但是任何错误的操作都可能返回nil。

如：Excel.NewWorkBook()，是返回workbook对象，错误时会返回的workbook将会是nil

## 安装

```bash
go get -u -v github.com/jinzhongmin/excel
```

## 例子

```go
package main

import (
	"github.com/jinzhongmin/excel"
)

func main() {
	Excel := excel.NewExcel()
	Const := excel.NewConst()
	Excel.Visible(true)

	workbook := Excel.NewWorkBook()
	sheet1 := workbook.Sheet(0)
	sheet1.Cell(1, 1).SetValue("Hello excel")
	sheet1.Cell(1, 1).GetFont().SetColor("FF00FF")
	byName := workbook.Sheet(sheet1.GetName())
	byName.Cell(1, 2).SetValue("get sheet by name")

	workbook.SaveAs("excel.xlsx", Const.FILEFORMATxlWorkbookDefault)

	Excel.Close()
}

```

## License

MIT
