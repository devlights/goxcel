package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"log"
)

var (
	_releaser = NewReleaser()
)

type (
	Goxcel struct {
		excel *ole.IDispatch
	}

	ReleaseFunc func()
)

func NewGoxcel() (*Goxcel, ReleaseFunc, error) {
	g := new(Goxcel)

	err := g.init()

	g.Releaser().Add(func() error {
		_ = g.quit()
		_ = g.release()

		return nil
	})

	startReleaserFunc := func() {
		e := g.Releaser().Release()
		if e != nil {
			log.Println(e)
		}
	}

	return g, startReleaserFunc, err
}

func (g *Goxcel) init() error {
	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		return err
	}

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return err
	}

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}

	g.excel = excel

	return nil
}

func (g *Goxcel) quit() error {
	_, err := oleutil.CallMethod(g.ComObject(), "Quit")
	if err != nil {
		log.Println(err)
	}

	return nil
}

func (g *Goxcel) release() error {
	g.excel.Release()
	ole.CoUninitialize()

	return nil
}

func (g *Goxcel) Goxcel() *Goxcel {
	return g
}

func (g *Goxcel) ComObject() *ole.IDispatch {
	return g.excel
}

func (g *Goxcel) Releaser() *Releaser {
	return _releaser
}

func (g *Goxcel) SetDisplayAlerts(value bool) error {
	_, err := oleutil.PutProperty(g.ComObject(), "DisplayAlerts", value)
	return err
}

func (g *Goxcel) SetVisible(value bool) error {
	_, err := oleutil.PutProperty(g.ComObject(), "Visible", value)
	return err
}

func (g *Goxcel) Workbooks() (*Workbooks, error) {
	wb, err := oleutil.GetProperty(g.ComObject(), "Workbooks")
	if err != nil {
		return nil, err
	}

	workbooks := NewWorkbooks(g, wb.ToIDispatch())

	return workbooks, nil
}

func (g *Goxcel) ActiveWindow() (*Window, error) {
	w, err := oleutil.GetProperty(g.ComObject(), "ActiveWindow")
	if err != nil {
		return nil, err
	}

	window := NewWindow(g, w.ToIDispatch())

	return window, nil
}

func (g *Goxcel) ActiveWorkbook() (*Workbook, ReleaseFunc, error) {
	wbs, err := g.Workbooks()
	if err != nil {
		return nil, nil, err
	}

	w, err := oleutil.GetProperty(g.ComObject(), "ActiveWorkbook")
	if err != nil {
		return nil, nil, err
	}

	workbook, wbReleaseFn := NewWorkbook(wbs, w.ToIDispatch())

	return workbook, wbReleaseFn, nil
}

func (g *Goxcel) ActiveSheet() (*Worksheet, *Workbook, ReleaseFunc, error) {
	wb, wbReleaseFn, err := g.ActiveWorkbook()
	if err != nil {
		return nil, nil, nil, err
	}

	s, err := oleutil.GetProperty(g.ComObject(), "ActiveSheet")
	if err != nil {
		return nil, nil, nil, err
	}

	sheet := NewWorksheet(wb, s.ToIDispatch())

	return sheet, wb, wbReleaseFn, nil
}
