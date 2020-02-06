package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Cell struct {
		ws *Worksheet
		c  *ole.IDispatch
	}
)

func NewCell(ws *Worksheet, c *ole.IDispatch) *Cell {
	return &Cell{ws: ws, c: c}
}

func (c *Cell) Value() (interface{}, error) {
	v, err := oleutil.GetProperty(c.c, "Value")
	if err != nil {
		return nil, err
	}

	return v.Value(), nil
}

func (c *Cell) SetValue(value interface{}) error {
	_, err := oleutil.PutProperty(c.c, "Value", value)
	if err != nil {
		return err
	}

	return nil
}
