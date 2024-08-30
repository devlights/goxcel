package goxcel

import (
	"fmt"
	"github.com/devlights/goxcel/constants"
	"github.com/devlights/goxcel/testutil"
	"testing"
)

func TestXlRange_Find(t *testing.T) {
	quit := MustInitGoxcel()
	defer quit()

	excel, release := MustNewGoxcel()
	defer release()

	wbs := excel.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	rows := []int{1, 2, 3, 4, 5}
	for _, row := range rows {
		cols := []int{1, 2, 3, 4, 5}
		for _, col := range cols {
			c, _ := ws.Cells(row, col)
			c.MustSetValue(fmt.Sprintf("%v,%v", row, col))
		}
	}

	//
	// Range.Find
	//
	rng, _ := ws.UsedRange()
	after, _ := rng.Cells(1, 1)
	foundRange, found, err := rng.Find("1,", after)
	if err != nil {
		t.Fatal(err)
	}

	if !found {
		t.Fatalf("expected range to be found, got nothing")
	}

	value, err := foundRange.Value()
	if err != nil {
		t.Fatal(err)
	}

	t.Log(value)

	//
	// Range.FindNext
	//
	after, _ = foundRange.Cells(1, 1)
	foundRange, found, err = rng.FindNext(after)
	if err != nil {
		t.Fatal(err)
	}

	if !found {
		t.Fatalf("expected range to be found, got nothing 2")
	}

	value, err = foundRange.Value()
	if err != nil {
		t.Fatal(err)
	}

	t.Log(value)

	//
	// Range.FindPrevious
	//
	before, _ := foundRange.Cells(1, 1)
	foundRange, found, err = rng.FindPrevious(before)
	if err != nil {
		t.Fatal(err)
	}

	if !found {
		t.Fatalf("expected range to be found, got nothing 3")
	}

	value, err = foundRange.Value()
	if err != nil {
		t.Fatal(err)
	}

	t.Log(value)
}

func TestXlRange_CopyPicture(t *testing.T) {
	quit := MustInitGoxcel()
	defer quit()

	excel, release := MustNewGoxcel()
	defer release()

	wbs := excel.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	rows := []int{1, 2, 3, 4, 5}
	for _, row := range rows {
		cols := []int{1, 2, 3, 4, 5}
		for _, col := range cols {
			c, _ := ws.Cells(row, col)
			c.MustSetValue(fmt.Sprintf("%v,%v", row, col))
		}
	}

	rng, _ := ws.UsedRange()
	_ = rng.Select()

	err := rng.CopyPicture(constants.XlScreen, constants.XlBitmap)
	if err != nil {
		t.Error(err)
	}

	// 結果はクリップボードにコピーされている
}

func TestXlRange_PageBreak(t *testing.T) {
	quit := MustInitGoxcel()
	defer quit()

	excel, release := MustNewGoxcel()
	defer release()

	wbs := excel.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	rows := []int{2, 4, 6, 8, 10}
	for _, row := range rows {
		ra, _ := ws.Range(row, 5, row, 5)

		err := ra.PageBreak(constants.XlPageBreakManual)
		if err != nil {
			t.Error(err)
		}

		ra.MustSetValue("hello")
	}

	hpb, _ := ws.HPageBreaks()
	if count, _ := hpb.Count(); int(count) != len(rows) {
		t.Errorf("[want] 5\t[got] %v", count)
	}
}

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
