package goxcel

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	XlRange struct {
		ws *Worksheet
		r  *ole.IDispatch
	}
)

func NewRange(ws *Worksheet, r *ole.IDispatch) *XlRange {
	xlrange := &XlRange{
		ws: ws,
		r:  r,
	}

	xlrange.Releaser().Add(func() error {
		xlrange.ComObject().Release()
		return nil
	})

	return xlrange
}

func (r *XlRange) ComObject() *ole.IDispatch {
	return r.r
}

func (r *XlRange) Goxcel() *Goxcel {
	return r.ws.wb.wbs.g
}

func (r *XlRange) Releaser() *Releaser {
	return r.Goxcel().Releaser()
}

func (r *XlRange) Cells(row int, col int) (*Cell, error) {
	if row <= 0 {
		e := fmt.Errorf("%w [row]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	if col <= 0 {
		e := fmt.Errorf("%w [col]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	c, err := oleutil.GetProperty(r.ComObject(), "Cells", row, col)
	if err != nil {
		return nil, err
	}

	cell := NewCell(r.ws, c.ToIDispatch())

	return cell, nil
}
