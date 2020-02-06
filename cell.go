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
	cell := &Cell{
		ws: ws,
		c:  c,
	}

	releaser.Add(func() error {
		cell.c.Release()
		return nil
	})

	return cell
}

func (c *Cell) ComObject() *ole.IDispatch {
	return c.c
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
