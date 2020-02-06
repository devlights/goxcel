package goxcel

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Worksheet struct {
		wb *Workbook
		ws *ole.IDispatch
	}
)

func NewWorksheet(wb *Workbook, ws *ole.IDispatch) *Worksheet {
	s := &Worksheet{
		wb: wb,
		ws: ws,
	}

	releaser.Add(func() error {
		s.ws.Release()
		return nil
	})

	return s
}

func (ws *Worksheet) ComObject() *ole.IDispatch {
	return ws.ws
}

func (ws *Worksheet) Name() (string, error) {
	v, err := oleutil.GetProperty(ws.ComObject(), "Name")
	if err != nil {
		return "", err
	}

	name := v.Value().(string)
	return name, nil
}

func (ws *Worksheet) Activate() error {
	_, err := oleutil.CallMethod(ws.ComObject(), "Activate")
	return err
}

func (ws *Worksheet) Cells(row int, col int) (*Cell, error) {
	if row <= 0 {
		e := fmt.Errorf("%w [row]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	if col <= 0 {
		e := fmt.Errorf("%w [col]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	c, err := oleutil.GetProperty(ws.ComObject(), "Cells", row, col)
	if err != nil {
		return nil, err
	}

	cell := NewCell(ws, c.ToIDispatch())

	return cell, nil
}
