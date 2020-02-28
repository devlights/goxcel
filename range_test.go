package goxcel

import (
	"github.com/devlights/goxcel/testutil"
	"testing"
)

func TestXlRange_Count(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()
	if err != nil {
		t.Error(err)
	}

	defer r()

	wbs, _ := g.Workbooks()

	wb, wbReleaseFn, _ := wbs.Add()
	defer wbReleaseFn()

	ws, _ := wb.Sheets(1)

	xlRange, err := ws.Range(1, 1, 2, 2)
	if err != nil {
		t.Error(err)
	}

	count, err := xlRange.Count()
	if err != nil {
		t.Error(err)
	}

	i := int(count)
	if i != 4 {
		t.Errorf("want: 4\tgot: %d", i)
	}
}

func TestXlRange_Clear(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	// Arrange
	g, r, _ := NewGoxcel()
	defer r()

	wbs, _ := g.Workbooks()
	wb, wbr, _ := wbs.Add()
	defer wbr()

	ws, _ := wb.Sheets(1)
	ra, _ := ws.Range(1, 1, 1, 1)

	_, _ = ra.Walk(func(xlsRange *XlRange, xlsCell *Cell, row, col int) error {
		_ = xlsCell.SetValue("helloworld")
		return nil
	})

	_, _ = ra.Walk(func(xlsRange *XlRange, xlsCell *Cell, row, col int) error {
		if s, _ := xlsCell.String(); s != "helloworld" {
			t.Errorf("want: helloworld\tgot: %v", s)
		}

		return nil
	})

	// Act
	err := ra.Clear()
	if err != nil {
		t.Error(err)
	}

	// Assert
	_, _ = ra.Walk(func(xlsRange *XlRange, xlsCell *Cell, row, col int) error {
		if s, _ := xlsCell.String(); s != "" {
			t.Errorf("want: empty\tgot: %v", s)
		}

		return nil
	})
}

func TestRange_Font(t *testing.T) {
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
	ra, _ := ws.Range(1, 1, 1, 1)
	_, _ = ra.Walk(func(xlsRange *XlRange, xlsCell *Cell, row, col int) error {
		_ = xlsCell.SetValue("helloworld")
		return nil
	})

	// Act
	f, err := ra.Font()
	if err != nil {
		t.Error(err)
	}

	if f == nil {
		t.Errorf("want: not nil\tgot nil")
	}
}

func TestRange_Interior(t *testing.T) {
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
	ra, _ := ws.Range(1, 1, 1, 1)
	_, _ = ra.Walk(func(xlsRange *XlRange, xlsCell *Cell, row, col int) error {
		_ = xlsCell.SetValue("helloworld")
		return nil
	})

	// Act
	interior, err := ra.Interior()
	if err != nil {
		t.Error(err)
	}

	if interior == nil {
		t.Errorf("want: not nil\tgot nil")
	}
}
