package goxcel

import (
	"fmt"

	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Worksheet struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newWorksheet(goxcelObj GoxcelObject, comObj *ole.IDispatch) *Worksheet {
	s := &Worksheet{
		goxcelObj: goxcelObj,
		comObj:    comObj,
	}

	s.Releaser().Add(func() error {
		s.ComObject().Release()
		return nil
	})

	return s
}

func NewWorksheet(wb *Workbook, ws *ole.IDispatch) *Worksheet {
	return newWorksheet(wb, ws)
}

func NewWorksheetFromWorksheets(ws *Worksheets, comObj *ole.IDispatch) *Worksheet {
	return newWorksheet(ws, comObj)
}

func (ws *Worksheet) ComObject() *ole.IDispatch {
	return ws.comObj
}

func (ws *Worksheet) Goxcel() *Goxcel {
	return ws.goxcelObj.Goxcel()
}

func (ws *Worksheet) Releaser() *Releaser {
	return ws.Goxcel().Releaser()
}

func (ws *Worksheet) Name() (string, error) {
	v, err := oleutil.GetProperty(ws.ComObject(), "Name")
	if err != nil {
		return "", err
	}

	name := v.Value().(string)
	return name, nil
}

func (ws *Worksheet) SetName(name string) error {
	_, err := oleutil.PutProperty(ws.ComObject(), "Name", name)
	return err
}

func (ws *Worksheet) Activate() error {
	_, err := oleutil.CallMethod(ws.ComObject(), "Activate")
	return err
}

func (ws *Worksheet) Range(startRow, startCol, endRow, endCol int) (*XlRange, error) {
	startCell, err := ws.Cells(startRow, startCol)
	if err != nil {
		return nil, err
	}

	endCell, err := ws.Cells(endRow, endCol)
	if err != nil {
		return nil, err
	}

	v, err := oleutil.GetProperty(ws.ComObject(), "Range", startCell.ComObject(), endCell.ComObject())
	if err != nil {
		return nil, err
	}

	newrange := NewRange(ws, v.ToIDispatch())
	return newrange, nil
}

func (ws *Worksheet) AllRange() (*XlRange, error) {
	ra, err := oleutil.GetProperty(ws.ComObject(), "Cells")
	if err != nil {
		return nil, err
	}

	xlrange := NewRange(ws, ra.ToIDispatch())

	return xlrange, nil
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

	c, err := oleutil.GetProperty(ws.ComObject(), "Cells", row, col)
	if err != nil {
		return nil, err
	}

	cell := NewCell(ws, c.ToIDispatch())

	return cell, nil
}

func (ws *Worksheet) MustCells(row int, col int) *Cell {
	c, err := ws.Cells(row, col)
	if err != nil {
		panic(err)
	}

	return c
}

func (ws *Worksheet) PageSetup() (*PageSetup, error) {
	p, err := oleutil.GetProperty(ws.ComObject(), "PageSetup")
	if err != nil {
		return nil, err
	}

	pagesetup := NewPageSetup(ws, p.ToIDispatch())

	return pagesetup, nil
}

func (ws *Worksheet) HPageBreaks() (*HPageBreaks, error) {
	v, err := oleutil.GetProperty(ws.ComObject(), "HPageBreaks")
	if err != nil {
		return nil, err
	}

	hpbs := NewHPageBreaks(ws, v.ToIDispatch())

	return hpbs, nil
}

func (ws *Worksheet) VPageBreaks() (*VPageBreaks, error) {
	v, err := oleutil.GetProperty(ws.ComObject(), "VPageBreaks")
	if err != nil {
		return nil, err
	}

	vpbs := NewVPageBreaks(ws, v.ToIDispatch())

	return vpbs, nil
}

func (ws *Worksheet) Shapes() (*Shapes, error) {
	v, err := oleutil.GetProperty(ws.ComObject(), "Shapes")
	if err != nil {
		return nil, err
	}

	shapes := NewShapesFromWorksheet(ws, v.ToIDispatch())

	return shapes, nil
}

func (ws *Worksheet) CopySheet(dest *Worksheet, after bool) error {
	var e error
	if after {
		_, e = oleutil.CallMethod(ws.ComObject(), "Copy", nil, dest.ComObject())
	} else {
		_, e = oleutil.CallMethod(ws.ComObject(), "Copy", dest.ComObject())
	}

	return e
}

func (ws *Worksheet) ExportAsFixedFormat(fmtType constants.XlFixedFormatType, path string) error {
	_, err := oleutil.CallMethod(ws.ComObject(), "ExportAsFixedFormat", int(fmtType), path)
	return err
}
