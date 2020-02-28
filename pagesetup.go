package goxcel

import (
	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	PageSetup struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func NewPageSetup(ws *Worksheet, ps *ole.IDispatch) *PageSetup {
	p := &PageSetup{
		comObj:    ps,
		goxcelObj: ws,
	}

	p.Releaser().Add(func() error {
		p.ComObject().Release()
		return nil
	})

	return p
}

func (p *PageSetup) ComObject() *ole.IDispatch {
	return p.comObj
}

func (p *PageSetup) Goxcel() *Goxcel {
	return p.goxcelObj.Goxcel()
}

func (p *PageSetup) Releaser() *Releaser {
	return p.Goxcel().Releaser()
}

func (p *PageSetup) SetOrientation(value constants.XlPageOrientation) error {
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
