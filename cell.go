package goxcel

import (
	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Cell struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newCell(goxcelObj GoxcelObject, comObj *ole.IDispatch) *Cell {
	cell := &Cell{
		goxcelObj: goxcelObj,
		comObj:    comObj,
	}

	cell.Releaser().Add(func() error {
		cell.ComObject().Release()
		return nil
	})

	return cell
}

func NewCell(ws *Worksheet, c *ole.IDispatch) *Cell {
	return newCell(ws, c)
}

func NewCellFromRange(ra *XlRange, c *ole.IDispatch) *Cell {
	return newCell(ra, c)
}

func (c *Cell) ComObject() *ole.IDispatch {
	return c.comObj
}

func (c *Cell) Goxcel() *Goxcel {
	return c.goxcelObj.Goxcel()
}

func (c *Cell) Releaser() *Releaser {
	return c.Goxcel().Releaser()
}

func (c *Cell) Value() (interface{}, error) {
	v, err := oleutil.GetProperty(c.ComObject(), "Value")
	if err != nil {
		return nil, err
	}

	return v.Value(), nil
}

func (c *Cell) SetValue(value interface{}) error {
	_, err := oleutil.PutProperty(c.ComObject(), "Value", value)
	if err != nil {
		return err
	}

	return nil
}

func (c *Cell) Select() error {
	_, err := oleutil.CallMethod(c.ComObject(), "Select")
	return err
}

func (c *Cell) String() (string, error) {
	value, err := c.Value()
	if err != nil {
		return "", err
	}

	s, ok := value.(string)
	if !ok {
		return "", ValueCantConvertToString
	}

	return s, nil
}

func (c *Cell) Font() (*Font, error) {
	v, err := oleutil.GetProperty(c.ComObject(), "Font")
	if err != nil {
		return nil, err
	}

	font := NewFont(c, v.ToIDispatch())

	return font, nil
}

func (c *Cell) Interior() (*Interior, error) {
	v, err := oleutil.GetProperty(c.ComObject(), "Interior")
	if err != nil {
		return nil, err
	}

	interior := NewInterior(c, v.ToIDispatch())

	return interior, nil
}

func (c *Cell) SetNumberFormatLocal(format constants.NumberFormatLocal) error {
	_, err := oleutil.PutProperty(c.ComObject(), "NumberFormatLocal", string(format))
	if err != nil {
		return err
	}

	return nil
}
