package goxcel

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Worksheets struct {
		wss *ole.IDispatch
		wb  *Workbook
	}
)

func NewWorkSheets(wb *Workbook, wss *ole.IDispatch) *Worksheets {
	w := &Worksheets{
		wss: wss,
		wb:  wb,
	}

	releaser.Add(func() error {
		w.ComObject().Release()
		return nil
	})

	return w
}

func (w *Worksheets) ComObject() *ole.IDispatch {
	return w.wss
}

func (w *Worksheets) Count() (int64, error) {
	count, err := oleutil.GetProperty(w.ComObject(), "Count")
	if err != nil {
		return 0, err
	}

	return count.Val, nil
}

func (w *Worksheets) Item(index int) (*Worksheet, error) {
	if index <= 0 {
		e := fmt.Errorf("%w [index]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	ws, err := oleutil.GetProperty(w.ComObject(), "Item", index)
	if err != nil {
		return nil, err
	}

	worksheet := NewWorksheet(w.wb, ws.ToIDispatch())

	return worksheet, nil
}

func (w *Worksheets) Add() (*Worksheet, error) {
	ws, err := oleutil.CallMethod(w.ComObject(), "Add")
	if err != nil {
		return nil, err
	}

	worksheet := NewWorksheet(w.wb, ws.ToIDispatch())

	return worksheet, nil
}
