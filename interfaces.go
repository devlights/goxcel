package goxcel

import "github.com/go-ole/go-ole"

type (
	HasComObject interface {
		ComObject() *ole.IDispatch
	}
)
