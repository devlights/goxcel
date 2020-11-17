package goxcel

import (
	"testing"

	"github.com/devlights/goxcel/testutil"
)

func TestWorksheet_CopySheet(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, _ := NewGoxcel()
	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(false)

	wbs, _ := g.Workbooks()
	wb1, wbr1, _ := wbs.Add()
	defer wbr1()
	wb2, wbr2, _ := wbs.Add()
	defer wbr2()

	ws1, _ := wb1.Sheets(1)
	c, _ := ws1.Cells(1, 1)
	_ = c.SetValue("worksheet.Copy")

	ws2, _ := wb2.Sheets(1)

	for _, beforeAfter := range []bool{false, true} {
		if err := ws1.CopySheet(ws2, beforeAfter); err != nil {
			t.Errorf("ws1.CopySheet(ws2, %v) got %v", beforeAfter, err)
		}
	}

	_ = g.SetVisible(true)

	for i := 0; i < 3; i++ {
		testutil.Interval()
	}
}

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
