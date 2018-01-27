package excel

import (
	"log"
	"strconv"

	ole "github.com/go-ole/go-ole"
)

//Col ..
type Col struct {
	obj *ole.IDispatch
}

//Cell ..
func (c *Col) Cell(index int) *Cell {
	obj, err := c.obj.GetProperty("Cells", 1, index)
	if err != nil {
		log.Println("in func Cell:", err)
		return nil
	}
	cell := new(Cell)
	cell.obj = obj.ToIDispatch()
	return cell
}

//Len ..
func (c *Col) Len() int {
	obj, err := c.obj.GetProperty("Rows")
	if err != nil {
		log.Println("in func Len:", err)
		return 0
	}
	rows := obj.ToIDispatch()

	obj, err = rows.GetProperty("Count")
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
func (c *Col) Range() *Range {
	obj, err := c.obj.GetProperty("Range", "A1:A"+strconv.Itoa(c.Len()))
	if err != nil {
		log.Println("in func Range:", err)
		return nil
	}
	r := new(Range)
	r.obj = obj.ToIDispatch()
	return r
}
