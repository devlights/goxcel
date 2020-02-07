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

func (w *Workbook) ComObject() *ole.IDispatch {
	return w.wb
}

func (w *Workbook) Goxcel() *Goxcel {
	return w.wbs.g
}

func (w *Workbook) WorkSheets() (*Worksheets, error) {
	wss, err := oleutil.GetProperty(w.ComObject(), "Sheets")
	if err != nil {
		return nil, err
	}

	worksheets := NewWorkSheets(w, wss.ToIDispatch())

	return worksheets, nil
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

func (w *Workbook) Save() error {
	_, err := oleutil.CallMethod(w.ComObject(), "Save")
	return err
}

func (w *Workbook) SaveAs(filePath string) error {
	return w.SaveAsWithFileFormat(filePath, XlFileFormatXlOpenXMLWorkbook)
}

func (w *Workbook) SaveAsWithFileFormat(filePath string, format XlFileFormat) error {
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
