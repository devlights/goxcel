package util

import (
	"github.com/devlights/goxcel"
	"github.com/go-ole/go-ole"
)

type (
	ComObjectProcFunc func(c *ole.IDispatch) (interface{}, *ole.IDispatch, error)
)

func ProcessComObject(com goxcel.ComReleaser, fn ComObjectProcFunc) (interface{}, *ole.IDispatch, error) {
	v, newComObj, err := fn(com.ComObject())
	if err != nil {
		return nil, nil, err
	}

	if newComObj != nil {
		com.Releaser().Add(func() error {
			newComObj.Release()
			return nil
		})
	}

	return v, newComObj, nil
}
