package excel

import (
	"log"
	"path/filepath"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

//Excel ..
type Excel struct {
	obj *ole.IDispatch
}

//NewExcel ..
func NewExcel() *Excel {
	ole.CoInitialize(0)

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		log.Fatalln("in func NewExcel:", err)
	}
	obj, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatalln("in func NewExcel:", err)
	}
	e := new(Excel)
	e.obj = obj

	if _, err := oleutil.PutProperty(obj, "Visible", false); err != nil {
		log.Println("in func NewExcel:", err)
	}
	if _, err := oleutil.PutProperty(obj, "DisplayAlerts", false); err != nil {
		log.Println("in func NewExcel:", err)
	}

	return e
}

func (e *Excel) workbooks() *ole.IDispatch {
	obj, err := e.obj.GetProperty("Workbooks")
	if err != nil {
		log.Println(err)
		e.Close()
	}
	return obj.ToIDispatch()
}

//NewWorkBook ..
func (e *Excel) NewWorkBook() *Workbook {
	obj, err := e.workbooks().CallMethod("Add")
	if err != nil {
		log.Println("in func NewWorkBook:", err)
		return nil
	}
	workbook := new(Workbook)
	workbook.obj = obj.ToIDispatch()

	return workbook
}

//OpenWorkBook ..
func (e *Excel) OpenWorkBook(file string) *Workbook {
	path, err := filepath.Abs(file)
	if err != nil {
		log.Println("in func OpenWorkBook:", err)
		return nil
	}

	obj, err := e.workbooks().CallMethod("Open", path)
	if err != nil {
		log.Println("in func OpenWorkBook:", err)
		return nil
	}
	workbook := new(Workbook)
	workbook.obj = obj.ToIDispatch()

	return workbook
}

//Close ..
func (e *Excel) Close() {
	wbs := e.WorkBooks()
	for i := range wbs {
		wbs[i].Close()
	}

	e.obj.CallMethod("Quit")
	e.obj.Release()
	ole.CoUninitialize()
}

//Visible ..
func (e *Excel) Visible(v bool) {
	if _, err := oleutil.PutProperty(e.obj, "Visible", v); err != nil {
		log.Println("in func Visible:", err)
	}
}

//Alert ..
func (e *Excel) Alert(v bool) {
	if _, err := oleutil.PutProperty(e.obj, "DisplayAlerts", v); err != nil {
		log.Println("in func Alert:", err)
	}
}

//WorkBooks ..
func (e *Excel) WorkBooks() []*Workbook {
	workbooksObj := e.workbooks()
	obj, err := workbooksObj.GetProperty("Count")
	if err != nil {
		log.Println("in func WorkBooks:", err)
	}
	len := int(obj.Val)

	workbooks := make([]*Workbook, 0)
	for i := 0; i < len; i++ {
		if obj, err := workbooksObj.GetProperty("Item", i+1); err != nil {
			log.Println("in func WorkBooks:", err)
		} else {
			workbook := new(Workbook)
			workbook.obj = obj.ToIDispatch()
			workbooks = append(workbooks, workbook)
		}

	}

	return workbooks
}
