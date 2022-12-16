package goxcel

import (
	"github.com/devlights/goxcel/constants"
	"github.com/devlights/goxcel/testutil"
	"testing"
)

// TestFont_Misc is tested function following cases
//
// - Bold
//
// - Italic
//
// - Underline
//
// - Strikethrough
//
// - Size
//
// - Name
//
// - Color
func TestFont_Misc(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	// Arrange
	g, r, _ := NewGoxcel()
	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(false)

	wbs, _ := g.Workbooks()
	wb, wbReleaseFn, _ := wbs.Add()
	defer wbReleaseFn()

	ws, _ := wb.Sheets(1)
	c, _ := ws.Cells(1, 1)
	_ = c.SetValue("helloworld")

	f, _ := c.Font()

	// Pre assert
	if isBold, err := f.Bold(); isBold || err != nil {
		t.Errorf("want not bold\tgot bold")
	}

	// Act
	err := f.SetBold(true)
	if err != nil {
		t.Error(err)
	}

	err = f.SetItalic(true)
	if err != nil {
		t.Error(err)
	}

	err = f.SetUnderline(constants.XlUnderlineStyleDouble)
	if err != nil {
		t.Error(err)
	}

	err = f.SetStrikethrough(true)
	if err != nil {
		t.Error(err)
	}

	err = f.SetSize(40)
	if err != nil {
		t.Error(err)
	}

	err = f.SetName("ＭＳ ゴシック")
	if err != nil {
		t.Error(err)
	}

	err = f.SetColor(constants.RgbRed)
	if err != nil {
		t.Error(err)
	}

	_ = g.SetVisible(true)
	testutil.Interval()
	testutil.Interval()
}
