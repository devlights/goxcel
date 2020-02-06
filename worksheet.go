package goxcel

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Worksheet struct {
		b *Workbook
		s *ole.IDispatch
	}
)

func NewWorksheet(b *Workbook, s *ole.IDispatch) *Worksheet {
	return &Worksheet{b: b, s: s}
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

	c, err := oleutil.GetProperty(ws.s, "Cells", row, col)
	if err != nil {
		return nil, err
	}

	cell := NewCell(ws, c.ToIDispatch())

	return cell, nil
}
