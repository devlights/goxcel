package goxcel

import (
	"fmt"

	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Workbook struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func NewWorkbook(wbs *Workbooks, wb *ole.IDispatch) (*Workbook, ReleaseFunc) {
	b := &Workbook{
		goxcelObj: wbs,
		comObj:    wb,
	}

	b.Releaser().Add(func() error {
		b.ComObject().Release()
		return nil
	})

	r := func() {
		_ = b.SetSaved(true)
		_ = b.Close()
	}

	return b, r
}

func (w *Workbook) ComObject() *ole.IDispatch {
	return w.comObj
}

func (w *Workbook) Goxcel() *Goxcel {
	return w.goxcelObj.Goxcel()
}

func (w *Workbook) Releaser() *Releaser {
	return w.Goxcel().Releaser()
}

func (w *Workbook) WorkSheets() (*Worksheets, error) {
	wss, err := oleutil.GetProperty(w.ComObject(), "Sheets")
	if err != nil {
		return nil, err
	}

	worksheets := NewWorkSheets(w, wss.ToIDispatch())

	return worksheets, nil
}

func (w *Workbook) MustWorkSheets() *Worksheets {
	wss, err := w.WorkSheets()
	if err != nil {
		panic(err)
	}

	return wss
}

func (w *Workbook) Sheets(index int) (*Worksheet, error) {
	if index <= 0 {
		e := fmt.Errorf("%w [index]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	s, err := oleutil.GetProperty(w.ComObject(), "Worksheets", index)
	if err != nil {
		return nil, err
	}

	ws := NewWorksheet(w, s.ToIDispatch())

	return ws, nil
}

func (w *Workbook) MustSheets(index int) *Worksheet {
	ws, err := w.Sheets(index)
	if err != nil {
		panic(err)
	}

	return ws
}

func (w *Workbook) Save() error {
	_, err := oleutil.CallMethod(w.ComObject(), "Save")
	return err
}

func (w *Workbook) SaveAs(filePath string) error {
	return w.SaveAsWithFileFormat(filePath, constants.XlOpenXMLWorkbook)
}

func (w *Workbook) SaveAsWithFileFormat(filePath string, format constants.XlFileFormat) error {
	_, err := oleutil.CallMethod(w.ComObject(), "SaveAs", filePath, int(format))
	return err
}

func (w *Workbook) SetSaved(value bool) error {
	_, err := oleutil.PutProperty(w.ComObject(), "Saved", value)
	return err
}

func (w *Workbook) Close() error {
	_, err := oleutil.CallMethod(w.ComObject(), "Close", false)
	return err
}

func (w *Workbook) PrintOut() error {
	_, err := oleutil.CallMethod(w.ComObject(), "PrintOut", nil)
	return err
}

func (w *Workbook) ExportAsFixedFormat(fmtType constants.XlFixedFormatType, path string) error {
	_, err := oleutil.CallMethod(w.ComObject(), "ExportAsFixedFormat", int(fmtType), path)
	return err
}
