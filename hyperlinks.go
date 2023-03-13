package goxcel

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	HyperLinks struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newHyperLinks(goxcelObj GoxcelObject, comObj *ole.IDispatch) *HyperLinks {
	hl := new(HyperLinks)

	hl.goxcelObj = goxcelObj
	hl.comObj = comObj

	hl.Releaser().Add(func() error {
		hl.ComObject().Release()
		return nil
	})

	return hl
}

func NewHyperLinks(ws *Worksheet, comObj *ole.IDispatch) *HyperLinks {
	return newHyperLinks(ws, comObj)
}

func NewHyperLinksFromHyperLinks(hl *HyperLinks, comObj *ole.IDispatch) *HyperLinks {
	return newHyperLinks(hl, comObj)
}

func (hl *HyperLinks) Goxcel() *Goxcel {
	return hl.goxcelObj.Goxcel()
}

func (hl *HyperLinks) Releaser() *Releaser {
	return hl.Goxcel().Releaser()
}

func (hl *HyperLinks) ComObject() *ole.IDispatch {
	return hl.comObj
}

func (hl *HyperLinks) Add(ra *XlRange, address string, subAddress string, screenTip string, textToDisplay string) error {
	_, err := oleutil.CallMethod(hl.ComObject(), "Add", ra.ComObject(), address, subAddress, screenTip, textToDisplay)
	if err != nil {
		return err
	}

	return nil
}

func (hl *HyperLinks) Item(index int) (*HyperLinks, error) {
	if index <= 0 {
		e := fmt.Errorf("%w [index]", ValueMustBeGreaterThanZero)
		return nil, e
	}

	item, err := oleutil.GetProperty(hl.ComObject(), "Item", index)
	if err != nil {
		return nil, err
	}

	hl2 := NewHyperLinksFromHyperLinks(hl, item.ToIDispatch())

	return hl2, nil
}

func (hl *HyperLinks) Count() (int64, error) {
	count, err := oleutil.GetProperty(hl.ComObject(), "Count")
	if err != nil {
		return 0, err
	}

	return count.Val, nil
}

func (hl *HyperLinks) Delete() error {
	_, err := oleutil.CallMethod(hl.ComObject(), "Delete")
	if err != nil {
		return err
	}

	return nil
}
