package goxcel

import (
	"errors"
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

var (
	SkipRow = errors.New("skip row")
	SkipCol = errors.New("skip col")
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

func (r *XlRange) Clear() error {
	_, err := oleutil.CallMethod(r.ComObject(), "Clear")
	if err != nil {
		return err
	}

	return nil
}

func (r *XlRange) Count() (int64, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Count")
	if err != nil {
		return 0, err
	}

	count := v.Val

	return count, nil
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

func (r *XlRange) Columns() (*XlRange, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Columns")
	if err != nil {
		return nil, err
	}

	xlrange := NewRange(r.ws, v.ToIDispatch())
	return xlrange, nil
}

func (r *XlRange) Rows() (*XlRange, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Rows")
	if err != nil {
		return nil, err
	}

	xlrange := NewRange(r.ws, v.ToIDispatch())
	return xlrange, nil
}

func (r *XlRange) Walk(walkFn func(r *XlRange, c *Cell, row, col int) error) (*Cell, error) {
	rows, err := r.Rows()
	if err != nil {
		return nil, err
	}

	cols, err := r.Columns()
	if err != nil {
		return nil, err
	}

	rowCount, err := rows.Count()
	if err != nil {
		return nil, err
	}

	colCount, err := cols.Count()
	if err != nil {
		return nil, err
	}

	for rowIndex := 1; rowIndex <= int(rowCount); rowIndex++ {
		for colIndex := 1; colIndex <= int(colCount); colIndex++ {
			cell, err := r.Cells(rowIndex, colIndex)
			if err != nil {
				return cell, err
			}

			err = walkFn(r, cell, rowIndex, colIndex)
			if err != nil {
				if errors.Is(err, SkipCol) {
					continue
				}

				if errors.Is(err, SkipRow) {
					break
				}

				return cell, err
			}
		}
	}

	return nil, nil
}
