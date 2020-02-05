package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"log"
)

type Mode int

const (
	ModeNew Mode = iota
	ModeOpen
)

type (
	Goxcel struct {
		Args      *Args
		excel     *ole.IDispatch
		workbooks *Workbooks
	}

	Args struct {
		FilePath string
		FileMode Mode
	}

	ReleaseFunc func(withQuit bool)
)

func NewArgs(filePath string) *Args {
	a := new(Args)
	a.FilePath = filePath

	a.FileMode = ModeOpen
	if a.FilePath == "" {
		a.FileMode = ModeNew
	}

	return a
}

func NewGoxcel(args *Args) (*Goxcel, ReleaseFunc, error) {
	g := new(Goxcel)
	g.Args = args

	err := g.init()

	return g, g.release, err
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

func (g *Goxcel) release(withQuit bool) {
	defer ole.CoUninitialize()
	defer g.excel.Release()

	if withQuit {
		defer func() {
			err := g.Quit()
			log.Fatal(err)
		}()
	}
}

func (g *Goxcel) Visible(value bool) error {
	_, err := oleutil.PutProperty(g.excel, "Visible", value)
	return err
}

func (g *Goxcel) Workbooks() (*Workbooks, error) {
	if g.workbooks != nil {
		return g.workbooks, nil
	}

	wb, err := oleutil.GetProperty(g.excel, "Workbooks")
	if err != nil {
		return nil, err
	}

	g.workbooks = NewWorkbooks(g, wb.ToIDispatch())

	return g.workbooks, nil
}

func (g *Goxcel) Quit() error {
	_, err := oleutil.CallMethod(g.excel, "Quit")
	if err != nil {
		return err
	}

	return nil
}
