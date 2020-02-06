package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Workbooks struct {
		g  *Goxcel
		wb *ole.IDispatch
	}
)

func NewWorkbooks(g *Goxcel, wb *ole.IDispatch) *Workbooks {
	w := &Workbooks{
		g:  g,
		wb: wb,
	}

	releaser.Add(func() error {
		w.wb.Release()
		return nil
	})

	return w
}

func (w *Workbooks) Add() (*Workbook, error) {
	wb, err := oleutil.CallMethod(w.wb, "Add", nil)
	if err != nil {
		return nil, err
	}

	workbook := NewWorkbook(w, wb.ToIDispatch())

	return workbook, nil
}

func (w *Workbooks) Open(filePath string) (*Workbook, error) {
	wb, err := oleutil.CallMethod(w.wb, "Open", filePath)
	if err != nil {
		return nil, err
	}

	workbook := NewWorkbook(w, wb.ToIDispatch())

	return workbook, nil
}
