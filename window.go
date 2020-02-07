package goxcel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Window struct {
		g *Goxcel
		w *ole.IDispatch
	}
)

func NewWindow(g *Goxcel, w *ole.IDispatch) *Window {
	win := &Window{
		g: g,
		w: w,
	}

	releaser.Add(func() error {
		win.w.Release()
		return nil
	})

	return win
}

func (w *Window) ComObject() *ole.IDispatch {
	return w.w
}

func (w *Window) Goxcel() *Goxcel {
	return w.g
}

func (w *Window) SetZoom(zoomRate int) error {
	_, err := oleutil.PutProperty(w.ComObject(), "Zoom", zoomRate)
	return err
}
