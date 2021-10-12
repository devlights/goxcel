package goxcel

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/devlights/goxcel/testutil"
)

func TestWorkbooks_MustMethods(t *testing.T) {
	f := MustInitGoxcel()
	defer f()

	g, r := MustNewGoxcel()
	defer r()

	g.MustSilent(true)

	wbs := g.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	if err := wb.SetSaved(true); err != nil {
		t.Error(err)
	}

	if err := wb.Close(); err != nil {
		t.Error(err)
	}
}

func TestWorkbooks_Add(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)

	wb, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	b, _, err := wb.Add()
	if err != nil {
		t.Error(err)
	}

	testutil.Interval()

	err = b.SetSaved(true)
	if err != nil {
		t.Error(err)
	}

	err = b.Close()
	if err != nil {
		t.Error(err)
	}
}

func TestWorkbooks_Open(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)

	wb, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	userHomeDir, _ := os.UserHomeDir()
	xlsxPath := filepath.Join(userHomeDir, "Book1.xlsx")
	b, _, err := wb.Open(xlsxPath)
	if err != nil {
		t.Error(err)
	}

	testutil.Interval()

	err = b.SetSaved(true)
	if err != nil {
		t.Error(err)
	}

	err = b.Close()
	if err != nil {
		t.Error(err)
	}
}
