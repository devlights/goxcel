package goxcel

import "github.com/go-ole/go-ole"

//noinspection GoNameStartsWithPackageName
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

	ComReleaser interface {
		HasReleaser
		HasComObject
	}

	GoxcelObject interface {
		HasGoxcel
		ComReleaser
	}
)
