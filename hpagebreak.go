package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	HPageBreak struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newHPageBreak(goxcelObj GoxcelObject, comObj *ole.IDispatch) *HPageBreak {
	hpb := new(HPageBreak)

	hpb.goxcelObj = goxcelObj
	hpb.comObj = comObj

	hpb.Releaser().Add(func() error {
		hpb.ComObject().Release()
		return nil
	})

	return hpb
}

func NewHPageBreak(hpbs *HPageBreaks, comObj *ole.IDispatch) *HPageBreak {
	return newHPageBreak(hpbs, comObj)
}

func (hpb *HPageBreak) Goxcel() *Goxcel {
	return hpb.goxcelObj.Goxcel()
}

func (hpb *HPageBreak) Releaser() *Releaser {
	return hpb.Goxcel().Releaser()
}

func (hpb *HPageBreak) ComObject() *ole.IDispatch {
	return hpb.comObj
}

func (hpb *HPageBreak) Location() (*XlRange, error) {
	v, err := oleutil.GetProperty(hpb.ComObject(), "Location")
	if err != nil {
		return nil, err
	}

	ra := NewRangeFromHPageBreak(hpb, v.ToIDispatch())

	return ra, nil
}

func (hpb *HPageBreak) SetLocation(ra *XlRange) error {
	_, err := oleutil.PutProperty(hpb.ComObject(), "Location", ra.comObj)
	if err != nil {
		return err
	}

	return nil
}
