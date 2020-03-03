package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Shapes struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newShapes(goxcelObj GoxcelObject, comObj *ole.IDispatch) *Shapes {
	shapes := new(Shapes)

	shapes.goxcelObj = goxcelObj
	shapes.comObj = comObj

	return shapes
}

func NewShapesFromWorksheet(ws *Worksheet, comObj *ole.IDispatch) *Shapes {
	return newShapes(ws, comObj)
}

func (ss *Shapes) Goxcel() *Goxcel {
	return ss.goxcelObj.Goxcel()
}

func (ss *Shapes) Releaser() *Releaser {
	return ss.Goxcel().Releaser()
}

func (ss *Shapes) ComObject() *ole.IDispatch {
	return ss.comObj
}

func (ss *Shapes) Count() (int32, error) {
	v, err := oleutil.GetProperty(ss.ComObject(), "Count")
	if err != nil {
		return 0, err
	}

	count, ok := v.Value().(int32)
	if !ok {
		return 0, ValueCantConvertToInt
	}

	return count, nil
}

func (ss *Shapes) Item(index int) (*Shape, error) {
	v, err := oleutil.CallMethod(ss.ComObject(), "Item", index)
	if err != nil {
		return nil, err
	}

	shape := NewShape(ss, v.ToIDispatch())

	return shape, nil
}

func (ss *Shapes) Walk(walkFn func(s *Shape, index int) error) (*Shape, error) {
	count, err := ss.Count()
	if err != nil {
		return nil, err
	}

	for i := 1; i <= int(count); i++ {
		s, err := ss.Item(i)
		if err != nil {
			return nil, err
		}

		err = walkFn(s, i)
		if err != nil {
			return s, err
		}
	}

	return nil, nil
}
