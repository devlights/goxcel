package goxcel

import (
	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Shape struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newShape(goxcelObj GoxcelObject, comObj *ole.IDispatch) *Shape {
	shape := new(Shape)

	shape.goxcelObj = goxcelObj
	shape.comObj = comObj

	shape.Releaser().Add(func() error {
		shape.ComObject().Release()
		return nil
	})

	return shape
}

func NewShape(ss *Shapes, c *ole.IDispatch) *Shape {
	return newShape(ss, c)
}

func (s *Shape) Goxcel() *Goxcel {
	return s.goxcelObj.Goxcel()
}

func (s *Shape) Releaser() *Releaser {
	return s.Goxcel().Releaser()
}

func (s *Shape) ComObject() *ole.IDispatch {
	return s.comObj
}

func (s *Shape) TopLeftCell() (*XlRange, error) {
	v, err := oleutil.GetProperty(s.ComObject(), "TopLeftCell")
	if err != nil {
		return nil, err
	}

	ra := NewRangeFromShape(s, v.ToIDispatch())

	return ra, nil
}

func (s *Shape) Type() (constants.MsoShapeType, error) {
	v, err := oleutil.GetProperty(s.ComObject(), "Type")
	if err != nil {
		return constants.MsoShapeType(0), err
	}

	shapeType, ok := v.Value().(int)
	if !ok {
		return constants.MsoShapeType(0), ValueCantConvertToInt
	}

	return constants.MsoShapeType(shapeType), nil
}
