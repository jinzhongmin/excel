package excel

import (
	"log"

	ole "github.com/go-ole/go-ole"
)

//Row ..
type Row struct {
	obj *ole.IDispatch
}

//Cell ..
func (r *Row) Cell(index int) *Cell {
	obj, err := r.obj.GetProperty("Cells", index, 1)
	if err != nil {
		log.Println("in func Cell:", err)
		return nil
	}
	cell := new(Cell)
	cell.obj = obj.ToIDispatch()

	return cell
}

//Len ..
func (r *Row) Len() int {
	obj, err := r.obj.GetProperty("Columns")
	if err != nil {
		log.Println("in func Len:", err)
		return 0
	}
	cols := obj.ToIDispatch()

	obj, err = cols.GetProperty("Count")
	if err != nil {
		log.Println("in func Len:", err)
		return 0
	}

	len, ok := obj.Value().(int32)
	if ok {
		return int(len)
	}
	return 0
}

//Range ..
func (r *Row) Range() *Range {
	obj, err := r.obj.GetProperty("Range", "A1:"+NumToLetter(r.Len(), 1))
	if err != nil {
		log.Println("in func Range", err)
		return nil
	}

	rg := new(Range)
	rg.obj = obj.ToIDispatch()
	return rg
}
