package excel

import (
	"log"

	ole "github.com/go-ole/go-ole"
)

//Sheet ..
type Sheet struct {
	obj *ole.IDispatch
}

//GetName ..
func (s *Sheet) GetName() string {
	obj, err := s.obj.GetProperty("Name")
	if err != nil {
		log.Println("in func GetName:", err)
		return ""
	}

	name := obj.Value()
	v, ok := name.(string)
	if ok {
		return v
	}
	return ""
}

//SetName ..
func (s *Sheet) SetName(name string) {
	if _, err := s.obj.PutProperty("Name", name); err != nil {
		log.Println("in func SetName:", err)
	}

	return
}

//Row ..
func (s *Sheet) Row(r int) *Row {
	return s.UsedRange().Row(r)
}

//Col ..
func (s *Sheet) Col(c int) *Col {
	return s.UsedRange().Col(c)
}

//Cell ..
func (s *Sheet) Cell(col int, row int) *Cell {
	obj, err := s.obj.GetProperty("Cells", row, col)
	if err != nil {
		log.Println("in func Cell:", err)
		return nil
	}

	cell := new(Cell)
	cell.obj = obj.ToIDispatch()
	return cell
}

//UsedRange ..
func (s *Sheet) UsedRange() *Range {
	obj, err := s.obj.GetProperty("UsedRange")
	if err != nil {
		log.Println("in func UsedRange:", err)
		return nil

	}

	usedRange := new(Range)
	usedRange.obj = obj.ToIDispatch()
	return usedRange
}
