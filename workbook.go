package excel

import (
	"log"
	"path/filepath"

	ole "github.com/go-ole/go-ole"
)

//Workbook ..
type Workbook struct {
	obj *ole.IDispatch
}

func (wb *Workbook) sheets() *ole.IDispatch {
	obj, err := wb.obj.GetProperty("Sheets")
	if err != nil {
		log.Println(err)
		return nil
	}

	return obj.ToIDispatch()
}

//Close ..
func (wb *Workbook) Close() {
	wb.obj.CallMethod("Colse")
	wb.obj.Release()
}

//AppendSheet ..
func (wb *Workbook) AppendSheet() *Sheet {
	sheets := wb.sheets()
	if sheets == nil {
		return nil
	}

	len := 0
	var lastSheet *ole.IDispatch
	if obj, err := sheets.GetProperty("Count"); err != nil {
		log.Println("in func AppendSheet:", err)

	} else {
		len = int(obj.Val)
	}
	if obj, err := sheets.GetProperty("Item", len); err != nil || len == 0 {
		log.Println("in func AppendSheet:", err)
	} else {
		lastSheet = obj.ToIDispatch()
	}

	if obj, err := sheets.CallMethod("Add", nil, lastSheet); err != nil {
		log.Println("in func AppendSheet:", err)
	} else {
		sheet := new(Sheet)
		sheet.obj = obj.ToIDispatch()
		return sheet
	}

	log.Println("in func AppendSheet: unknown err")
	return nil
}

//Sheet ..
func (wb *Workbook) Sheet(index interface{}) *Sheet {
	sheets := wb.sheets()
	if sheets == nil {
		return nil
	}

	obj, err := sheets.GetProperty("Count")
	if err != nil {
		log.Println("in func Sheet:", err)
		return nil
	}
	len := int(obj.Val)

	switch v := index.(type) {
	case string:
		sheet := new(Sheet)
		obj, err := sheets.GetProperty("Item", v)
		if err != nil {
			log.Println("in func Sheet:", err)
			return nil
		}
		sheet.obj = obj.ToIDispatch()
		return sheet

	case int:
		if v < len {
			sheet := new(Sheet)
			obj, err := sheets.GetProperty("Item", v+1)
			if err != nil {
				log.Println("in func Sheet:", err)
				return nil
			}
			sheet.obj = obj.ToIDispatch()
			return sheet
		}
		log.Println("in func Sheet: err index > len(sheets)")
		return nil
	}

	log.Println("in func Sheet: unknown err")
	return nil
}

//Sheets ..
func (wb *Workbook) Sheets() []*Sheet {
	_sheets := wb.sheets()
	if _sheets == nil {
		return nil
	}

	len := 0
	obj, err := _sheets.GetProperty("Count")
	if err != nil {
		log.Println("in func Sheets:", err)
		return nil
	}
	len = int(obj.Val)

	sheets := make([]*Sheet, 0)
	for i := 0; i < len; i++ {
		obj, err := _sheets.GetProperty("Item", i+1)
		if err == nil {
			sheet := new(Sheet)
			sheet.obj = obj.ToIDispatch()
			sheets = append(sheets, sheet)
			continue
		}
		log.Println("in func Sheets:", err)
	}

	return sheets
}

//Save ..
func (wb *Workbook) Save() {
	if _, err := wb.obj.CallMethod("save"); err != nil {
		log.Println("in func Save:", err)
	}
}

//SaveAs ..
func (wb *Workbook) SaveAs(name string, format int) {
	path, err := filepath.Abs(name)
	if err != nil {
		log.Println("in func SaveAs:", err)
		return
	}

	if _, err = wb.obj.CallMethod("SaveAs", path, format); err != nil {
		log.Println("in func SaveAs:", err)
		return
	}
}
