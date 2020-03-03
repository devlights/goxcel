package goxcel

import (
	"github.com/devlights/goxcel/testutil"
	"testing"
)

// TestWorksheet_Misc is test function following cases
//
// - HPageBreaks
func TestWorksheet_Misc(t *testing.T) {
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
	c, _ := ws.Cells(10, 1)
	_ = c.SetValue("helloworld")

	// Act
	hpbs, err := ws.HPageBreaks()
	if err != nil {
		t.Error(err)
	}

	// Assert
	if hpbs == nil {
		t.Errorf("want: not nil\tgot: nil")
	}

	_ = g.SetVisible(true)
	testutil.Interval()
	testutil.Interval()
}
