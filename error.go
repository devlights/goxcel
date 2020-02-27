package goxcel

import "errors"

var (
	ValueMustBeGreaterThanZero = errors.New("value must be greater than 0")
	ValueCantConvertToString   = errors.New("value can't convert to string type")
)
