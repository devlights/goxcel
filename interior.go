package goxcel

import (
	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Interior struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newInterior(goxcelObj GoxcelObject, comObj *ole.IDispatch) *Interior {
	i := new(Interior)

	i.goxcelObj = goxcelObj
	i.comObj = comObj

	i.Releaser().Add(func() error {
		i.comObj.Release()
		return nil
	})

	return i
}

func NewInterior(cell *Cell, comObj *ole.IDispatch) *Interior {
	return newInterior(cell, comObj)
}

func NewInteriorFromRange(ra *XlRange, comObj *ole.IDispatch) *Interior {
	return newInterior(ra, comObj)
}

func (i *Interior) Goxcel() *Goxcel {
	return i.goxcelObj.Goxcel()
}

func (i *Interior) Releaser() *Releaser {
	return i.Goxcel().Releaser()
}

func (i *Interior) ComObject() *ole.IDispatch {
	return i.comObj
}

func (i *Interior) Color() (constants.XlRgbColor, error) {
	v, err := oleutil.GetProperty(i.ComObject(), "Color")
	if err != nil {
		return constants.RgbBlack, err
	}

	color, ok := v.Value().(int)
	if !ok {
		return constants.RgbBlack, ValueCantConvertToInt
	}

	return constants.XlRgbColor(color), nil
}

func (i *Interior) SetColor(color constants.XlRgbColor) error {
	_, err := oleutil.PutProperty(i.ComObject(), "Color", int(color))
	if err != nil {
		return err
	}

	return nil
}
