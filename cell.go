package excel

import (
	"log"

	ole "github.com/go-ole/go-ole"
)

//Cell ..
type Cell struct {
	obj *ole.IDispatch
}

//GetValue ..
func (c *Cell) GetValue() string {
	obj, err := c.obj.GetProperty("Value")
	if err != nil {
		log.Println("in func GetValue:", err)
		return ""
	}
	val := obj.Value()
	v, ok := val.(string)
	if ok {
		return v
	}
	return ""
}

//SetValue ..
func (c *Cell) SetValue(val string) {
	if _, err := c.obj.PutProperty("Value", val); err != nil {
		log.Println("in func SetValue:", err)
	}
}

//GetFont ..
func (c *Cell) GetFont() *Font {
	obj, err := c.obj.GetProperty("Font")
	if err != nil {
		log.Println("in func GetFont:", err)
		return nil
	}
	font := new(Font)
	font.obj = obj.ToIDispatch()

	return font
}

//Range ..
func (c *Cell) Range() *Range {
	obj, err := c.obj.GetProperty("Range", "A1")
	if err != nil {
		log.Println("in func Range:", err)
		return nil
	}
	r := new(Range)
	r.obj = obj.ToIDispatch()

	return r
}
