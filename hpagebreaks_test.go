package goxcel

import (
	"github.com/devlights/goxcel/testutil"
	"testing"
)

// TestHPageBreaks_Misc is tested function following cases
//
// - Count
// - Item
func TestHPageBreaks_Misc(t *testing.T) {
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
	hpbs, err := ws.HPageBreaks()
	if err != nil {
		t.Error(err)
	}

	err = hpbs.Add(ra)
	if err != nil {
		t.Error(err)
	}

	// Assert
	count, err := hpbs.Count()
	if err != nil {
		t.Error(err)
	}

	if count != 1 {
		t.Errorf("want: 1\tgot: %v", count)
	}

	hpb, err := hpbs.Item(1)
	if err != nil {
		t.Error(err)
	}

	if hpb == nil {
		t.Errorf("want: not nil\tgot nil")
	}

	_ = g.SetVisible(true)
	testutil.Interval()
	testutil.Interval()
}
