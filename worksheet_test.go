package goxcel

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/devlights/goxcel/constants"
	"github.com/devlights/goxcel/testutil"
)

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

	const MAX_ROW_COUNT = 1048576
	if cnt != MAX_ROW_COUNT {
		t.Errorf("[want] %v\t[got] %v", MAX_ROW_COUNT, cnt)
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
