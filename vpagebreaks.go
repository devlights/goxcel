package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	VPageBreaks struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newVPageBreaks(goxcelObj GoxcelObject, comObj *ole.IDispatch) *VPageBreaks {
	vpbs := new(VPageBreaks)

	vpbs.goxcelObj = goxcelObj
	vpbs.comObj = comObj

	vpbs.Releaser().Add(func() error {
		vpbs.ComObject().Release()
		return nil
	})

	return vpbs
}

func NewVPageBreaks(ws *Worksheet, comObj *ole.IDispatch) *VPageBreaks {
	return newVPageBreaks(ws, comObj)
}

func (vpbs *VPageBreaks) Goxcel() *Goxcel {
	return vpbs.goxcelObj.Goxcel()
}

func (vpbs *VPageBreaks) Releaser() *Releaser {
	return vpbs.Goxcel().Releaser()
}

func (vpbs *VPageBreaks) ComObject() *ole.IDispatch {
	return vpbs.comObj
}

func (vpbs *VPageBreaks) Add(ra *XlRange) error {
	_, err := oleutil.CallMethod(vpbs.ComObject(), "Add", ra.ComObject())
	if err != nil {
		return err
	}

	return nil
}

func (vpbs *VPageBreaks) Count() (int32, error) {
	v, err := oleutil.GetProperty(vpbs.ComObject(), "Count")
	if err != nil {
		return 0, err
	}

	count, ok := v.Value().(int32)
	if !ok {
		return 0, ValueCantConvertToInt
	}

	return count, nil
}

func (vpbs *VPageBreaks) Item(index int) (*VPageBreak, error) {
	v, err := oleutil.GetProperty(vpbs.ComObject(), "Item", index)
	if err != nil {
		return nil, err
	}

	hpb := NewVPageBreak(vpbs, v.ToIDispatch())

	return hpb, nil
}

func (vpbs *VPageBreaks) Walk(walkFn func(hpb *VPageBreak) error) (*VPageBreak, error) {
	count, err := vpbs.Count()
	if err != nil {
		return nil, err
	}

	for i := 1; i <= int(count); i++ {
		hpb, err := vpbs.Item(i)
		if err != nil {
			return nil, err
		}

		err = walkFn(hpb)
		if err != nil {
			return hpb, err
		}
	}

	return nil, nil
}
