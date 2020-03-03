package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	VPageBreak struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newVPageBreak(goxcelObj GoxcelObject, comObj *ole.IDispatch) *VPageBreak {
	vpb := new(VPageBreak)

	vpb.goxcelObj = goxcelObj
	vpb.comObj = comObj

	vpb.Releaser().Add(func() error {
		vpb.ComObject().Release()
		return nil
	})

	return vpb
}

func NewVPageBreak(vpbs *VPageBreaks, comObj *ole.IDispatch) *VPageBreak {
	return newVPageBreak(vpbs, comObj)
}

func (vpb *VPageBreak) Goxcel() *Goxcel {
	return vpb.goxcelObj.Goxcel()
}

func (vpb *VPageBreak) Releaser() *Releaser {
	return vpb.Goxcel().Releaser()
}

func (vpb *VPageBreak) ComObject() *ole.IDispatch {
	return vpb.comObj
}

func (vpb *VPageBreak) Location() (*XlRange, error) {
	v, err := oleutil.GetProperty(vpb.ComObject(), "Location")
	if err != nil {
		return nil, err
	}

	ra := NewRangeFromVPageBreak(vpb, v.ToIDispatch())

	return ra, nil
}

func (vpb *VPageBreak) SetLocation(ra *XlRange) error {
	_, err := oleutil.PutProperty(vpb.ComObject(), "Location", ra.ComObject())
	if err != nil {
		return err
	}

	return nil
}
