package goxcel

import (
	"github.com/devlights/goxcel/testutil"
	"testing"
)

// TestHPageBreaks_Misc is tested function following cases
//
// - Location
// - SetLocation
func TestHPageBreak_Misc(t *testing.T) {
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

	ra, _ := ws.Range(10, 1, 10, 1)

	// Act
	hpbs, _ := ws.HPageBreaks()
	_ = hpbs.Add(ra)

	// Assert
	hpb, _ := hpbs.Item(1)
	location, err := hpb.Location()
	if err != nil {
		t.Error(err)
	}

	row, _ := location.Row()
	column, _ := location.Column()
	if row != 10 {
		t.Errorf("want: 10\tgot: %v", row)
	}

	if column != 1 {
		t.Errorf("want: 1\tgot %v", column)
	}

	_ = g.SetVisible(true)
	testutil.Interval()
	testutil.Interval()
}
