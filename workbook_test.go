package goxcel

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/devlights/goxcel/testutil"
)

func TestWorkbook_MustMethods(t *testing.T) {
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

	if err := wb.SetSaved(true); err != nil {
		t.Error(err)
	}

	if err := wb.Close(); err != nil {
		t.Error(err)
	}
}

func TestWorkbook_Save(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)

	wbs, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	userHomeDir, _ := os.UserHomeDir()
	xlsxPath := filepath.Join(userHomeDir, "Book1.xlsx")
	wb, wbReleaseFn, _ := wbs.Open(xlsxPath)
	defer wbReleaseFn()

	ws, _ := wb.Sheets(1)
	c, _ := ws.Cells(1, 1)
	_ = c.SetValue("こんにちわ goxcel")

	err = wb.Save()
	if err != nil {
		t.Error(err)
	}

	testutil.Interval()
}

func TestWorkbook_SaveAs(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)

	wbs, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	userHomeDir, _ := os.UserHomeDir()
	srcXlsxPath := filepath.Join(userHomeDir, "Book1.xlsx")
	wb, wbReleaseFn, _ := wbs.Open(srcXlsxPath)
	defer wbReleaseFn()

	ws, _ := wb.Sheets(1)
	c, _ := ws.Cells(1, 1)
	_ = c.SetValue("hello goxcel")

	dstXlsxPath := filepath.Join(userHomeDir, "Book2.xlsx")
	err = wb.SaveAs(dstXlsxPath)
	if err != nil {
		t.Error(err)
	}

	testutil.Interval()
}
