package goxcel

import (
	"github.com/devlights/goxcel/testutil"
	"testing"
)

func TestCell_Value(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()
	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)

	wbs, _ := g.Workbooks()
	wb, wbReleaseFn, _ := wbs.Add()
	defer wbReleaseFn()

	ws, err := wb.Sheets(1)
	if err != nil {
		t.Error(err)
	}

	testutil.Interval()

	c, err := ws.Cells(1, 1)
	if err != nil {
		t.Error(err)
	}

	err = c.SetValue("helloworld")
	if err != nil {
		t.Error(err)
	}

	v, err := c.Value()
	if err != nil {
		t.Error(err)
	}

	if v.(string) != "helloworld" {
		t.Errorf("Want: helloworld\tGot: %v", v)
	}

	testutil.Interval()

	err = wb.SetSaved(true)
	if err != nil {
		t.Error(err)
	}

	err = wb.Close()
	if err != nil {
		t.Error(err)
	}
}
