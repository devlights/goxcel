package goxcel

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Worksheets struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func NewWorkSheets(wb *Workbook, wss *ole.IDispatch) *Worksheets {
	w := &Worksheets{
		comObj:    wss,
		goxcelObj: wb,
	}

	w.Releaser().Add(func() error {
		w.ComObject().Release()
		return nil
	})

	return w
}

func (w *Worksheets) ComObject() *ole.IDispatch {
	return w.comObj
}

func (w *Worksheets) Goxcel() *Goxcel {
	return w.goxcelObj.Goxcel()
}

func (w *Worksheets) Releaser() *Releaser {
	return w.Goxcel().Releaser()
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

	worksheet := NewWorksheetFromWorksheets(w, ws.ToIDispatch())

	return worksheet, nil
}

func (w *Worksheets) Add() (*Worksheet, error) {
	ws, err := oleutil.CallMethod(w.ComObject(), "Add")
	if err != nil {
		return nil, err
	}

	worksheet := NewWorksheetFromWorksheets(w, ws.ToIDispatch())

	return worksheet, nil
}

func (w *Worksheets) Walk(walkFn func(ws *Worksheet, index int) error) (*Worksheet, error) {
	c, err := w.Count()
	if err != nil {
		return nil, err
	}

	count := int(c)
	for i := 0; i < count; i++ {
		ws, err := w.Item(i + 1)
		if err != nil {
			return nil, err
		}

		err = walkFn(ws, i)
		if err != nil {
			return ws, err
		}
	}

	return nil, nil
}
