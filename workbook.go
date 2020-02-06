package goxcel

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Workbook struct {
		wbs *Workbooks
		wb  *ole.IDispatch
	}
)

func NewWorkbook(wbs *Workbooks, wb *ole.IDispatch) *Workbook {
	b := &Workbook{
		wbs: wbs,
		wb:  wb,
	}

	releaser.Add(func() error {
		b.wb.Release()
		return nil
	})

	return b
}

func (w *Workbook) Sheets(index int) (*Worksheet, error) {
	if index <= 0 {
		e := fmt.Errorf("%w [index]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	s, err := oleutil.GetProperty(w.wb, "Worksheets", index)
	if err != nil {
		return nil, err
	}

	ws := NewWorksheet(w, s.ToIDispatch())

	return ws, nil
}

func (w *Workbook) Saved(value bool) error {
	_, err := oleutil.PutProperty(w.wb, "Saved", value)
	return err
}

func (w *Workbook) Close() error {
	_, err := oleutil.CallMethod(w.wb, "Close", false)
	return err
}
