package goxcel

import (
	"github.com/devlights/goxcel/constants"
	"github.com/devlights/goxcel/testutil"
	"testing"
)

// TestInterior_Misc is test function following cases
//
// - Color
func TestInterior_Misc(t *testing.T) {
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

	// Act
	interior, err := c.Interior()
	if err != nil {
		t.Error(err)
	}

	err = interior.SetColor(constants.RgbRed)
	if err != nil {
		t.Error(err)
	}

	// Assert
	color, err := interior.Color()
	if err != nil {
		t.Error(err)
	}

	if color != constants.RgbRed {
		t.Errorf("want: red(%v)\tgot %v", constants.RgbRed, color)
	}

	_ = g.SetVisible(true)
	testutil.Interval()
	testutil.Interval()
}
