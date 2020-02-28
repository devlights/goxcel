package goxcel

import (
	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	// Font represents an Excel Font
	//
	// REFERENCES::
	//   - https://docs.microsoft.com/ja-jp/office/vba/api/excel.font(object)
	Font struct {
		goxcelObj GoxcelObject
		comObj    *ole.IDispatch
	}
)

func newFont(goxcelObj GoxcelObject, comObj *ole.IDispatch) *Font {
	f := &Font{
		goxcelObj: goxcelObj,
		comObj:    comObj,
	}

	f.Releaser().Add(func() error {
		f.comObj.Release()
		return nil
	})

	return f
}

func NewFont(c *Cell, comObj *ole.IDispatch) *Font {
	return newFont(c, comObj)
}

func NewFontFromRange(r *XlRange, comObj *ole.IDispatch) *Font {
	return newFont(r, comObj)
}

func (f *Font) Goxcel() *Goxcel {
	return f.goxcelObj.Goxcel()
}

func (f *Font) Releaser() *Releaser {
	return f.goxcelObj.Releaser()
}

func (f *Font) ComObject() *ole.IDispatch {
	return f.comObj
}

func (f *Font) Bold() (bool, error) {
	v, err := oleutil.GetProperty(f.ComObject(), "Bold")
	if err != nil {
		return false, err
	}

	b, ok := v.Value().(bool)
	if !ok {
		return false, ValueCantConvertToBool
	}

	return b, nil
}

func (f *Font) SetBold(isBold bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Bold", isBold)
	if err != nil {
		return err
	}

	return nil
}

func (f *Font) Italic() (bool, error) {
	v, err := oleutil.GetProperty(f.ComObject(), "Italic")
	if err != nil {
		return false, err
	}

	b, ok := v.Value().(bool)
	if !ok {
		return false, ValueCantConvertToBool
	}

	return b, nil
}

func (f *Font) SetItalic(isItalic bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Italic", isItalic)
	if err != nil {
		return err
	}

	return nil
}

func (f *Font) Underline() (constants.XlUnderlineStyle, error) {
	v, err := oleutil.GetProperty(f.ComObject(), "Underline")
	if err != nil {
		return constants.XlUnderlineStyleSingleAccounting, err
	}

	i, ok := v.Value().(int)
	if !ok {
		return constants.XlUnderlineStyleSingleAccounting, ValueCantConvertToInt
	}

	return constants.XlUnderlineStyle(i), nil
}

func (f *Font) SetUnderline(value constants.XlUnderlineStyle) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Underline", int(value))
	if err != nil {
		return err
	}

	return nil
}

func (f *Font) Strikethrough() (bool, error) {
	v, err := oleutil.GetProperty(f.ComObject(), "Strikethrough")
	if err != nil {
		return false, err
	}

	b, ok := v.Value().(bool)
	if !ok {
		return false, ValueCantConvertToBool
	}

	return b, nil
}

func (f *Font) SetStrikethrough(isStrikethrough bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Strikethrough", isStrikethrough)
	if err != nil {
		return err
	}

	return nil
}

func (f *Font) Size() (int, error) {
	v, err := oleutil.GetProperty(f.ComObject(), "Size")
	if err != nil {
		return 0, err
	}

	fontSize, ok := v.Value().(int)
	if !ok {
		return 0, ValueCantConvertToInt
	}

	return fontSize, nil
}

func (f *Font) SetSize(fontSize int) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Size", fontSize)
	if err != nil {
		return err
	}

	return nil
}

// Name gets the font name.
//
// REFERENCES::
//   - https://www.officepro.jp/excelvba/cell_font/index2.html
func (f *Font) Name() (string, error) {
	v, err := oleutil.GetProperty(f.ComObject(), "Name")
	if err != nil {
		return "", err
	}

	fontName, ok := v.Value().(string)
	if !ok {
		return "", ValueCantConvertToString
	}

	return fontName, nil
}

// SetName sets the font name.
//
// REFERENCES::
//   - https://www.officepro.jp/excelvba/cell_font/index2.html
func (f *Font) SetName(fontName string) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Name", fontName)
	if err != nil {
		return err
	}

	return nil
}

func (f *Font) Color() (constants.XlRgbColor, error) {
	v, err := oleutil.GetProperty(f.ComObject(), "Color")
	if err != nil {
		return constants.RgbBlack, err
	}

	color, ok := v.Value().(int)
	if !ok {
		return constants.RgbBlack, ValueCantConvertToInt
	}

	return constants.XlRgbColor(color), nil
}

func (f *Font) SetColor(color constants.XlRgbColor) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Color", int(color))
	if err != nil {
		return err
	}

	return nil
}
