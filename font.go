package excel

import (
	"log"
	"strconv"

	"github.com/go-ole/go-ole"
)

//Font ..
type Font struct {
	obj *ole.IDispatch
}

//GetName ..
func (f *Font) GetName() string {
	obj, err := f.obj.GetProperty("Name")
	if err != nil {
		log.Println("in func GetName:", err)
		return ""
	}
	return obj.ToString()
}

//SetName ..
func (f *Font) SetName(name string) {
	if _, err := f.obj.PutProperty("Name", name); err != nil {
		log.Println("in func SetName:", err)
	}
}

//SetSize ..
func (f *Font) SetSize(size float64) {
	if _, err := f.obj.PutProperty("Size", size); err != nil {
		log.Println("in func SetSize:", err)
	}
}

//GetSize ..
func (f *Font) GetSize() float64 {
	obj, err := f.obj.GetProperty("Size")
	if err != nil {
		log.Println("in func GetSize:", err)
		return 0
	}
	v, ok := obj.Value().(float64)
	if ok {
		return v
	}
	return 0
}

//SetColor ..
func (f *Font) SetColor(hex string) {
	hexInt, _ := strconv.ParseInt(hex, 16, 64)
	f.obj.PutProperty("Color", float64(hexInt))
}

//GetColor ..
func (f *Font) GetColor() string {
	obj, err := f.obj.GetProperty("Color")
	if err != nil {
		log.Println("in func GetColor:", err)
		return ""
	}
	v, ok := obj.Value().(float64)
	if ok {
		return strconv.FormatInt(int64(v), 16)
	}
	return ""
}
