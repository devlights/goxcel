package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	HPageBreaks struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newHPageBreaks(goxcelObj GoxcelObject, comObj *ole.IDispatch) *HPageBreaks {
	hpbs := new(HPageBreaks)

	hpbs.goxcelObj = goxcelObj
	hpbs.comObj = comObj

	hpbs.Releaser().Add(func() error {
		hpbs.ComObject().Release()
		return nil
	})

	return hpbs
}

func NewHPageBreaks(ws *Worksheet, comObj *ole.IDispatch) *HPageBreaks {
	return newHPageBreaks(ws, comObj)
}

func (hpbs *HPageBreaks) Goxcel() *Goxcel {
	return hpbs.goxcelObj.Goxcel()
}

func (hpbs *HPageBreaks) Releaser() *Releaser {
	return hpbs.Goxcel().Releaser()
}

func (hpbs *HPageBreaks) ComObject() *ole.IDispatch {
	return hpbs.comObj
}

func (hpbs *HPageBreaks) Add(ra *XlRange) error {
	_, err := oleutil.CallMethod(hpbs.ComObject(), "Add", ra.ComObject())
	if err != nil {
		return err
	}

	return nil
}

func (hpbs *HPageBreaks) Count() (int32, error) {
	v, err := oleutil.GetProperty(hpbs.ComObject(), "Count")
	if err != nil {
		return 0, err
	}

	count, ok := v.Value().(int32)
	if !ok {
		return 0, ValueCantConvertToInt
	}

	return count, nil
}

func (hpbs *HPageBreaks) Item(index int) (*HPageBreak, error) {
	v, err := oleutil.GetProperty(hpbs.ComObject(), "Item", index)
	if err != nil {
		return nil, err
	}

	hpb := NewHPageBreak(hpbs, v.ToIDispatch())

	return hpb, nil
}

func (hpbs *HPageBreaks) Walk(walkFn func(hpb *HPageBreak, index int) error) (*HPageBreak, error) {
	count, err := hpbs.Count()
	if err != nil {
		return nil, err
	}

	for i := 1; i <= int(count); i++ {
		hpb, err := hpbs.Item(i)
		if err != nil {
			return nil, err
		}

		err = walkFn(hpb, i)
		if err != nil {
			return hpb, err
		}
	}

	return nil, nil
}
