package goxcel

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Workbook struct {
		wb *Workbooks
		w  *ole.IDispatch
	}
)

func NewWorkbook(wb *Workbooks, w *ole.IDispatch) *Workbook {
	return &Workbook{
		wb: wb,
		w:  w,
	}
}

func (w *Workbook) Sheets(index int) (*Worksheet, error) {
	if index <= 0 {
		e := fmt.Errorf("%w [index]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	s, err := oleutil.GetProperty(w.w, "Worksheets", index)
	if err != nil {
		return nil, err
	}

	ws := NewWorksheet(w, s.ToIDispatch())

	return ws, nil
}

func (w *Workbook) Saved(value bool) error {
	_, err := oleutil.PutProperty(w.w, "Saved", value)
	return err
}

func (w *Workbook) Close() error {
	_, err := oleutil.CallMethod(w.w, "Close", false)
	return err
}
