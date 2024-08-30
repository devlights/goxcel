package goxcel

import (
	"errors"
	"fmt"
	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/skanehira/clipboard-image/v2"
	"io"
)

type (
	XlRange struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

var (
	SkipRow = errors.New("skip row")
	SkipCol = errors.New("skip col")
)

func newRange(goxcelObj GoxcelObject, comObj *ole.IDispatch) *XlRange {
	xlrange := &XlRange{
		goxcelObj: goxcelObj,
		comObj:    comObj,
	}

	xlrange.Releaser().Add(func() error {
		xlrange.ComObject().Release()
		return nil
	})

	return xlrange
}

func NewRange(ws *Worksheet, r *ole.IDispatch) *XlRange {
	return newRange(ws, r)
}

func NewRangeFromCell(cell *Cell, c *ole.IDispatch) *XlRange {
	return newRange(cell, c)
}

func NewRangeFromRange(ra *XlRange, c *ole.IDispatch) *XlRange {
	return newRange(ra, c)
}

func NewRangeFromHPageBreak(hpb *HPageBreak, c *ole.IDispatch) *XlRange {
	return newRange(hpb, c)
}

func NewRangeFromVPageBreak(vpb *VPageBreak, c *ole.IDispatch) *XlRange {
	return newRange(vpb, c)
}

func NewRangeFromShape(s *Shape, c *ole.IDispatch) *XlRange {
	return newRange(s, c)
}

func NewRangeFromWorksheet(ws *Worksheet, c *ole.IDispatch) *XlRange {
	return newRange(ws, c)
}

func (r *XlRange) ComObject() *ole.IDispatch {
	return r.comObj
}

func (r *XlRange) Goxcel() *Goxcel {
	return r.goxcelObj.Goxcel()
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

func (r *XlRange) Count() (int32, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Count")
	if err != nil {
		return 0, err
	}

	count := v.Value().(int32)

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

	cell := NewCellFromRange(r, c.ToIDispatch())

	return cell, nil
}

func (r *XlRange) Font() (*Font, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Font")
	if err != nil {
		return nil, err
	}

	font := NewFontFromRange(r, v.ToIDispatch())

	return font, nil
}

func (r *XlRange) Interior() (*Interior, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Interior")
	if err != nil {
		return nil, err
	}

	interior := NewInteriorFromRange(r, v.ToIDispatch())

	return interior, nil
}

func (r *XlRange) Column() (int32, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Column")
	if err != nil {
		return 0, err
	}

	column, ok := v.Value().(int32)
	if !ok {
		return 0, ValueCantConvertToInt
	}

	return column, nil
}

func (r *XlRange) Row() (int32, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Row")
	if err != nil {
		return 0, err
	}

	row, ok := v.Value().(int32)
	if !ok {
		return 0, ValueCantConvertToInt
	}

	return row, nil
}

func (r *XlRange) Columns() (*XlRange, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Columns")
	if err != nil {
		return nil, err
	}

	xlrange := NewRangeFromRange(r, v.ToIDispatch())
	return xlrange, nil
}

func (r *XlRange) Rows() (*XlRange, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Rows")
	if err != nil {
		return nil, err
	}

	xlrange := NewRangeFromRange(r, v.ToIDispatch())
	return xlrange, nil
}

func (r *XlRange) Value() (interface{}, error) {
	v, err := oleutil.GetProperty(r.ComObject(), "Value")
	if err != nil {
		return nil, err
	}

	return v.Value(), nil
}

func (r *XlRange) SetValue(value interface{}) error {
	_, err := oleutil.PutProperty(r.ComObject(), "Value", value)
	if err != nil {
		return err
	}

	return nil
}

func (r *XlRange) MustSetValue(value any) {
	err := r.SetValue(value)
	if err != nil {
		panic(err)
	}
}

func (r *XlRange) SetNumberFormatLocal(format constants.NumberFormatLocal) error {
	_, err := oleutil.PutProperty(r.ComObject(), "NumberFormatLocal", string(format))
	if err != nil {
		return err
	}

	return nil
}

func (r *XlRange) PageBreak(pageBreakType constants.XlPageBreak) error {
	_, err := oleutil.PutProperty(r.ComObject(), "PageBreak", int(pageBreakType))
	if err != nil {
		return err
	}

	return nil
}

func (r *XlRange) Select() error {
	_, err := oleutil.CallMethod(r.ComObject(), "Select")
	return err
}

func (r *XlRange) CopyPicture(appearance constants.XlPictureAppearance, format constants.XlCopyPictureFormat) error {
	_, err := oleutil.CallMethod(r.ComObject(), "CopyPicture", int(appearance), int(format))
	if err != nil {
		return err
	}

	return nil
}

func (r *XlRange) CopyPictureToFile(writer io.Writer, appearance constants.XlPictureAppearance, format constants.XlCopyPictureFormat) error {
	err := r.CopyPicture(appearance, format)
	if err != nil {
		return err
	}

	reader, err := clipboard.Read()
	if err != nil {
		return err
	}

	_, err = io.Copy(writer, reader)
	if err != nil {
		return err
	}

	return nil
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

func (r *XlRange) Find(what string, after *Cell) (*XlRange, bool, error) {
	var (
		lookIn          = constants.XlFindLookInValues
		lookAt          = constants.XlLookAtPart
		searchOrder     = constants.XlSearchOrderByRows
		searchDirection = constants.XlSearchDirectionNext
		matchCase       = false
		matchByte       = true
	)
	result, err := oleutil.CallMethod(r.ComObject(), "Find", what, after.ComObject(), int32(lookIn), int32(lookAt), int32(searchOrder), int32(searchDirection), matchCase, matchByte)
	if err != nil {
		return nil, false, err
	}

	dispatch := result.ToIDispatch()
	if dispatch == nil {
		return nil, false, nil
	}

	newR := NewRangeFromRange(r, dispatch)
	return newR, true, nil
}

func (r *XlRange) FindNext(after *Cell) (*XlRange, bool, error) {
	result, err := oleutil.CallMethod(r.ComObject(), "FindNext", after.ComObject())
	if err != nil {
		return nil, false, err
	}

	dispatch := result.ToIDispatch()
	if dispatch == nil {
		return nil, false, nil
	}

	newR := NewRangeFromRange(r, dispatch)
	return newR, true, nil
}

func (r *XlRange) FindPrevious(before *Cell) (*XlRange, bool, error) {
	result, err := oleutil.CallMethod(r.ComObject(), "FindPrevious", before.ComObject())
	if err != nil {
		return nil, false, err
	}

	dispatch := result.ToIDispatch()
	if dispatch == nil {
		return nil, false, nil
	}

	newR := NewRangeFromRange(r, dispatch)
	return newR, true, nil
}
