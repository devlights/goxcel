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
	return &Workbooks{
		g:  g,
		wb: wb,
	}
}

func (w *Workbooks) Add() (*Workbook, error) {
	b, err := oleutil.CallMethod(w.wb, "Add", nil)
	if err != nil {
		return nil, err
	}

	workbook := NewWorkbook(w, b.ToIDispatch())

	return workbook, nil
}
