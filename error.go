package goxcel

import "errors"

var (
	ValueMustBeGreaterThanZero = errors.New("value must be greater than 0")
	ValueCantConvertToInt      = errors.New("value can't convert to int type")
	ValueCantConvertToString   = errors.New("value can't convert to string type")
	ValueCantConvertToBool     = errors.New("value can't convert to bool type")
)
