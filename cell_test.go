package goxcel

import (
	"reflect"
	"testing"
	"time"

	"github.com/devlights/goxcel/constants"
	"github.com/devlights/goxcel/testutil"
)

func TestCell_PageBreak(t *testing.T) {
	quit := MustInitGoxcel()
	defer quit()

	excel, release := MustNewGoxcel()
	defer release()

	_ = excel.SetVisible(true)

	wbs := excel.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	rows := []int{2, 4, 6, 8, 10}
	for _, row := range rows {
		cell, _ := ws.Cells(row, 5)

		err := cell.PageBreak(constants.XlPageBreakManual)
		if err != nil {
			t.Error(err)
		}

		cell.MustSetValue("hello")
	}

	hpb, _ := ws.HPageBreaks()
	if count, _ := hpb.Count(); int(count) != len(rows) {
		t.Errorf("[want] 5\t[got] %v", count)
	}
}

func TestCell_End(t *testing.T) {
	f := MustInitGoxcel()
	defer f()

	g, r := MustNewGoxcel()
	defer r()

	g.MustSilent(true)

	wbs := g.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	rows, _ := ws.Rows()
	cols, _ := ws.Columns()
	rowsCount, _ := rows.Count()
	colsCount, _ := cols.Count()

	rowsCells, _ := ws.Cells(int(rowsCount), 1)
	colsCells, _ := ws.Cells(1, int(colsCount))

	rrange, err := rowsCells.End(constants.XlUp)
	if err != nil {
		t.Error(err)
	}

	rrow, _ := rrange.Row()
	if rrow != 1 {
		t.Errorf("[want] 1\t[got] %v", rrow)
	}

	crange, err := colsCells.End(constants.XlToLeft)
	if err != nil {
		t.Error(err)
	}

	ccol, _ := crange.Column()
	if ccol != 1 {
		t.Errorf("[want] 1\t[got] %v", ccol)
	}
}

func TestCell_MustMethods(t *testing.T) {
	f := MustInitGoxcel()
	defer f()

	g, r := MustNewGoxcel()
	defer r()

	g.MustSilent(true)

	wbs := g.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)
	if ws == nil {
		t.Error("ws is nil")
	}

	cell := ws.MustCells(1, 1)
	if cell == nil {
		t.Error("cell is nil")
	}

	if err := cell.SetValue("helloworld"); err != nil {
		t.Error(err)
	}

	v := cell.MustValue()
	if v != "helloworld" {
		t.Errorf("want: helloworld\tgot: %v", v)
	}
}

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

func TestCell_String(t *testing.T) {
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

	v, err := c.String()
	if err != nil {
		t.Error(err)
	}

	if reflect.TypeOf(v) != reflect.TypeOf("") {
		t.Errorf("Want: string\tGot: %v", reflect.TypeOf(v))
	}

	if v != "helloworld" {
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

func TestCell_Font(t *testing.T) {
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
	f, err := c.Font()
	if err != nil {
		t.Error(err)
	}

	if f == nil {
		t.Errorf("want: not nil\tgot nil")
	}
}

func TestCell_Interior(t *testing.T) {
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

	if interior == nil {
		t.Errorf("want: not nil\tgot nil")
	}
}

func TestCell_SetNumberFormatLocal(t *testing.T) {
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
	err := c.SetNumberFormatLocal(constants.FormatString)
	if err != nil {
		t.Error(err)
	}

	_ = g.SetVisible(true)
	time.Sleep(10 * time.Second)
}
