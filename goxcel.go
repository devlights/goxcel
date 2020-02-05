package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Mode int

const (
	ModeNew Mode = iota
	ModeOpen
)

type (
	Goxcel struct {
		Args  *Args
		excel *ole.IDispatch
	}

	Args struct {
		FilePath string
		FileMode Mode
	}

	ReleaseFunc func()
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

func (g *Goxcel) release() {
	defer ole.CoUninitialize()
	defer g.excel.Release()
}

func (g *Goxcel) Visible(value bool) error {
	_, err := oleutil.PutProperty(g.excel, "Visible", value)
	return err
}

func (g *Goxcel) Quit() error {
	_, err := oleutil.CallMethod(g.excel, "Quit")
	if err != nil {
		return err
	}

	return nil
}
