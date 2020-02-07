package goxcel

import "github.com/go-ole/go-ole"

type (
	HasReleaser interface {
		Releaser() *Releaser
	}

	HasComObject interface {
		ComObject() *ole.IDispatch
	}

	HasGoxcel interface {
		Goxcel() *Goxcel
	}
)
