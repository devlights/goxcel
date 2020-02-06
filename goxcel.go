package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"log"
)

var (
	releaser = NewReleaser()
)

type (
	Goxcel struct {
		excel     *ole.IDispatch
		workbooks *Workbooks
	}

	ReleaseFunc func()
)

func NewGoxcel() (*Goxcel, ReleaseFunc, error) {
	g := new(Goxcel)

	err := g.init()

	releaser.Add(g.quit)
	releaser.Add(g.release)

	startReleaserFunc := func() {
		e := releaser.Release()
		if e != nil {
			log.Println(e)
		}
	}

	return g, startReleaserFunc, err
}

func (g *Goxcel) init() error {
	err := ole.CoInitialize(0)
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

func (g *Goxcel) ComObject() *ole.IDispatch {
	return g.excel
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
	if g.workbooks != nil {
		return g.workbooks, nil
	}

	wb, err := oleutil.GetProperty(g.ComObject(), "Workbooks")
	if err != nil {
		return nil, err
	}

	g.workbooks = NewWorkbooks(g, wb.ToIDispatch())

	return g.workbooks, nil
}

func (g *Goxcel) ActiveWindow() (*Window, error) {
	w, err := oleutil.GetProperty(g.ComObject(), "ActiveWindow")
	if err != nil {
		return nil, err
	}

	window := NewWindow(g, w.ToIDispatch())

	return window, nil
}

func (g *Goxcel) ActiveWorkbook() (*Workbook, error) {
	wbs, err := g.Workbooks()
	if err != nil {
		return nil, err
	}

	w, err := oleutil.GetProperty(g.ComObject(), "ActiveWorkbook")
	if err != nil {
		return nil, err
	}

	workbook := NewWorkbook(wbs, w.ToIDispatch())

	return workbook, nil
}

func (g *Goxcel) ActiveSheet() (*Worksheet, error) {
	wb, err := g.ActiveWorkbook()
	if err != nil {
		return nil, err
	}

	s, err := oleutil.GetProperty(g.ComObject(), "ActiveSheet")
	if err != nil {
		return nil, err
	}

	sheet := NewWorksheet(wb, s.ToIDispatch())

	return sheet, nil
}
