package goxcel

import (
	"fmt"
	"os"
	"path/filepath"
	"testing"

	"github.com/devlights/goxcel/constants"
	"github.com/devlights/goxcel/testutil"
)

func TestWorksheet_ResetAllPageBreaks(t *testing.T) {
	quit := MustInitGoxcel()
	defer quit()

	excel, release := MustNewGoxcel()
	defer release()

	wbs := excel.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	cell := ws.MustCells(10, 5)
	cell.MustSetValue("hello")
	_ = cell.PageBreak(constants.XlPageBreakManual)

	vpb, _ := ws.VPageBreaks()
	if count, _ := vpb.Count(); count != 1 {
		t.Errorf("[want] 1\t[got] %v", count)
	}

	if err := ws.ResetAllPageBreaks(); err != nil {
		t.Error(err)
	}

	if count, _ := vpb.Count(); count != 0 {
		t.Errorf("[want] 0\t[got] %v", count)
	}
}

func TestWorksheet_UsedRange(t *testing.T) {
	f := MustInitGoxcel()
	defer f()

	g, r := MustNewGoxcel()
	defer r()

	g.MustSilent(true)

	wbs := g.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	c := ws.MustCells(1, 1)
	c.MustSetValue(fmt.Sprintf("%v_%v", 1, 1))

	c = ws.MustCells(100, 1)
	c.MustSetValue(fmt.Sprintf("%v_%v", 100, 1))

	c = ws.MustCells(50, 100)
	c.MustSetValue(fmt.Sprintf("%v_%v", 50, 100))

	// UsedRange は、書式設定されているだけのセルも範囲に入る
	c = ws.MustCells(200, 1)
	interior, _ := c.Interior()
	_ = interior.SetColor(constants.RgbGreen)

	ra, err := ws.UsedRange()
	if err != nil {
		t.Error(err)
	}

	rows, _ := ra.Rows()
	count, _ := rows.Count()
	if count != 200 {
		t.Errorf("[want] 200\t[got] %v", count)
	}

	cols, _ := ra.Columns()
	count, _ = cols.Count()
	if count != 100 {
		t.Errorf("[want] 100\t[got] %v", count)
	}
}

func TestWorksheet_MaxRowCol(t *testing.T) {
	f := MustInitGoxcel()
	defer f()

	g, r := MustNewGoxcel()
	defer r()

	g.MustSilent(true)

	wbs := g.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	c := ws.MustCells(1, 100)
	c.MustSetValue("hello")

	c = ws.MustCells(100, 1)
	c.MustSetValue("world")

	// MaxRow, MaxCol, MaxRowCol は、書式設定されているだけのセルは範囲に入れない
	c = ws.MustCells(200, 1)
	interior, _ := c.Interior()
	_ = interior.SetColor(constants.RgbGreen)

	maxRow, maxCol, err := ws.MaxRowCol(1, 1)
	if err != nil {
		t.Error(err)
	}

	if maxRow != 100 {
		t.Errorf("[want] 100\t[got] %v", maxRow)
	}
	if maxCol != 100 {
		t.Errorf("[want] 100\t[got] %v", maxCol)
	}

	c = ws.MustCells(1, int(maxCol))
	v := c.MustValue()
	if v != "hello" {
		t.Errorf("[want] hello\t[got] %v", v)
	}

	c = ws.MustCells(int(maxRow), 1)
	v = c.MustValue()
	if v != "world" {
		t.Errorf("[want] world\t[got] %v", v)
	}
}

func TestWorksheet_MaxCol(t *testing.T) {
	f := MustInitGoxcel()
	defer f()

	g, r := MustNewGoxcel()
	defer r()

	g.MustSilent(true)

	wbs := g.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	c := ws.MustCells(1, 100)
	c.MustSetValue("hello")

	maxCol, err := ws.MaxCol(1)
	if err != nil {
		t.Error(err)
	}

	if maxCol != 100 {
		t.Errorf("[want] 100\t[got] %v", maxCol)
	}
}

func TestWorksheet_MaxRow(t *testing.T) {
	f := MustInitGoxcel()
	defer f()

	g, r := MustNewGoxcel()
	defer r()

	g.MustSilent(true)

	wbs := g.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	c := ws.MustCells(100, 1)
	c.MustSetValue("hello")

	maxRow, err := ws.MaxRow(1)
	if err != nil {
		t.Error(err)
	}

	if maxRow != 100 {
		t.Errorf("[want] 100\t[got] %v", maxRow)
	}
}

func TestWorksheet_Columns(t *testing.T) {
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

	cols, err := ws.Columns()
	if err != nil {
		t.Error(err)
	}

	if cols == nil {
		t.Error("[want] not-nil\t[got] nil")
	}

	cnt, err := cols.Count()
	if err != nil {
		t.Error(err)
	}

	const MaxColCount = 16384
	if cnt != MaxColCount {
		t.Errorf("[want] %v\t[got] %v", MaxColCount, cnt)
	}

	if err := wb.SetSaved(true); err != nil {
		t.Error(err)
	}

	if err := wb.Close(); err != nil {
		t.Error(err)
	}
}

func TestWorksheet_Rows(t *testing.T) {
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

	rows, err := ws.Rows()
	if err != nil {
		t.Error(err)
	}

	if rows == nil {
		t.Error("[want] not-nil\t[got] nil")
	}

	cnt, err := rows.Count()
	if err != nil {
		t.Error(err)
	}

	const MaxRowCount = 1048576
	if cnt != MaxRowCount {
		t.Errorf("[want] %v\t[got] %v", MaxRowCount, cnt)
	}

	if err := wb.SetSaved(true); err != nil {
		t.Error(err)
	}

	if err := wb.Close(); err != nil {
		t.Error(err)
	}
}

func TestWorksheet_MustMethods(t *testing.T) {
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

	if err := wb.SetSaved(true); err != nil {
		t.Error(err)
	}

	if err := wb.Close(); err != nil {
		t.Error(err)
	}
}

func TestWorksheet_Name(t *testing.T) {
	cases := []struct {
		name string
		in   string
		out  string
	}{
		{"ascii-sheet-name", "helloworld", "helloworld"},
		{"non-ascii-sheet-name", "テストシート", "テストシート"},
	}

	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			testutil.Interval()
			defer testutil.Interval()

			g, r, _ := NewGoxcel()
			defer r()

			_ = g.SetDisplayAlerts(false)
			_ = g.SetVisible(false)

			wbs, _ := g.Workbooks()
			wb, wbr, _ := wbs.Add()
			defer wbr()

			ws, _ := wb.Sheets(1)

			if err := ws.SetName(c.in); err != nil {
				t.Errorf("ws.SetName(%v) got %v", c.in, err)
			}

			name, err := ws.Name()
			if err != nil {
				t.Errorf("ws.Name() got %v", err)
			}

			if name != c.out {
				t.Errorf("worksheet name [%v] != [%v]", name, c.out)
			}

			_ = g.SetVisible(true)
			for i := 0; i < 2; i++ {
				testutil.Interval()
			}
		})
	}
}

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

func TestWorksheet_HPageBreaks(t *testing.T) {
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

func TestWorksheet_ExportAsFixedFormat(t *testing.T) {
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
	p := filepath.Join(os.TempDir(), "goxcel-ExportAsFixedFormat-test.pdf")
	t.Log(p)

	err := ws.ExportAsFixedFormat(constants.XlTypePDF, p)
	if err != nil {
		t.Error(err)
	}

	// Assert
	_, err = os.Stat(p)
	if err != nil {
		t.Errorf("[want] file exists\t[got] file not exists\t%s", p)
	}

	_ = g.SetVisible(true)
	testutil.Interval()
	testutil.Interval()
}
