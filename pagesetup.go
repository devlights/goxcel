package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	PageSetup struct {
		ws *Worksheet
		ps *ole.IDispatch
	}
)

func NewPageSetup(ws *Worksheet, ps *ole.IDispatch) *PageSetup {
	p := &PageSetup{
		ps: ps,
		ws: ws,
	}

	releaser.Add(func() error {
		p.ps.Release()
		return nil
	})

	return p
}

func (p *PageSetup) ComObject() *ole.IDispatch {
	return p.ps
}

func (p *PageSetup) Goxcel() *Goxcel {
	return p.ws.wb.wbs.g
}

func (p *PageSetup) SetOrientation(value XlPageOrientation) error {
	_, err := oleutil.PutProperty(p.ComObject(), "Orientation", int(value))
	return err
}

func (p *PageSetup) SetCenterHeader(value interface{}) error {
	_, err := oleutil.PutProperty(p.ComObject(), "CenterHeader", value)
	return err
}

func (p *PageSetup) SetCenterFooter(value interface{}) error {
	_, err := oleutil.PutProperty(p.ComObject(), "CenterFooter", value)
	return err
}
