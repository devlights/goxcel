package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Workbooks struct {
		g   *Goxcel
		wbs *ole.IDispatch
	}
)

func NewWorkbooks(g *Goxcel, wbs *ole.IDispatch) *Workbooks {
	w := &Workbooks{
		g:   g,
		wbs: wbs,
	}

	releaser.Add(func() error {
		w.wbs.Release()
		return nil
	})

	return w
}

func (w *Workbooks) ComObject() *ole.IDispatch {
	return w.wbs
}

func (w *Workbooks) Goxcel() *Goxcel {
	return w.g
}

func (w *Workbooks) Add() (*Workbook, error) {
	wb, err := oleutil.CallMethod(w.ComObject(), "Add", nil)
	if err != nil {
		return nil, err
	}

	workbook := NewWorkbook(w, wb.ToIDispatch())

	return workbook, nil
}

func (w *Workbooks) Open(filePath string) (*Workbook, error) {
	wb, err := oleutil.CallMethod(w.ComObject(), "Open", filePath)
	if err != nil {
		return nil, err
	}

	workbook := NewWorkbook(w, wb.ToIDispatch())

	return workbook, nil
}
