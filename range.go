package excel

import (
	"log"

	"github.com/go-ole/go-ole"
)

//Range ..
type Range struct {
	obj *ole.IDispatch
}

//Range ..
func (r *Range) Range(x, y, w, h int) *Range {
	obj, err := r.obj.GetProperty("Range", NumToLetter(x, y)+":"+NumToLetter(x+w-1, y+h-1))
	if err != nil {
		log.Println("in func Range:", err)
		return nil

	}

	rg := new(Range)
	rg.obj = obj.ToIDispatch()
	return rg
}

//Cell ..
func (r *Range) Cell(col int, row int) *Cell {
	obj, err := r.obj.GetProperty("Item", row, col)
	if err != nil {
		log.Println("in func Cell:", err)
		return nil

	}

	cell := new(Cell)
	cell.obj = obj.ToIDispatch()
	return cell
}

//Row ..
func (r *Range) Row(num int) *Row {
	obj, err := r.obj.GetProperty("Rows", num)
	if err != nil {
		log.Println("in func Row:", err)
		return nil
	}

	row := new(Row)
	row.obj = obj.ToIDispatch()
	return row
}

//RowRange ..
func (r *Range) RowRange(num int) *Range {
	obj, err := r.obj.GetProperty("Rows", num)
	if err != nil {
		log.Println("in func RowRange:", err)
		return nil
	}

	row := new(Range)
	row.obj = obj.ToIDispatch()
	return row
}

//Col ..
func (r *Range) Col(num int) *Col {
	obj, err := r.obj.GetProperty("Columns", num)
	if err != nil {
		log.Println("in func Col:", err)
		return nil
	}

	col := new(Col)
	col.obj = obj.ToIDispatch()
	return col
}

//ColRange ..
func (r *Range) ColRange(num int) *Range {
	obj, err := r.obj.GetProperty("Columns", num)
	if err != nil {
		log.Println("in func ColRange:", err)
		return nil
	}

	col := new(Range)
	col.obj = obj.ToIDispatch()
	return col
}

//Copy ..
func (r *Range) Copy(dest *Range) {
	if _, err := r.obj.CallMethod("Copy", dest.Cell(1, 1).obj); err != nil {
		log.Println("in func Copy:", err)
	}

}

//ClearFormats ..
func (r *Range) ClearFormats() {
	if _, err := r.obj.CallMethod("ClearFormats"); err != nil {
		log.Println("in func ClearFormats:", err)
	}
}
