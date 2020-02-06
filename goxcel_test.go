package goxcel

import (
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestGoxcelStartup(t *testing.T) {
	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	err = g.Visible(true)
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)
}

func TestGoxcelWorkbooksAdd(t *testing.T) {
	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.Visible(true)

	wb, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	b, err := wb.Add()
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	err = b.Saved(true)
	if err != nil {
		t.Error(err)
	}

	err = b.Close()
	if err != nil {
		t.Error(err)
	}
}

func TestGoxcelWorkbooksOpen(t *testing.T) {
	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.Visible(true)

	wb, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	userHomeDir, _ := os.UserHomeDir()
	xlsxPath := filepath.Join(userHomeDir, "Book1.xlsx")
	b, err := wb.Open(xlsxPath)
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	err = b.Saved(true)
	if err != nil {
		t.Error(err)
	}

	err = b.Close()
	if err != nil {
		t.Error(err)
	}
}

func TestGoxcelWorkbookSave(t *testing.T) {
	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.Visible(true)

	wbs, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	userHomeDir, _ := os.UserHomeDir()
	xlsxPath := filepath.Join(userHomeDir, "Book1.xlsx")
	wb, _ := wbs.Open(xlsxPath)

	ws, _ := wb.Sheets(1)
	c, _ := ws.Cells(1, 1)
	_ = c.SetValue("hello goxcel")

	err = wb.Save()
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	_ = wb.Saved(true)
	_ = wb.Close()
}

func TestGoxcelCellValue(t *testing.T) {
	g, r, err := NewGoxcel()
	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.Visible(true)
	wbs, _ := g.Workbooks()
	wb, _ := wbs.Add()

	ws, err := wb.Sheets(1)
	if err != nil {
		t.Error(err)
	}

	time.Sleep(2 * time.Second)

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

	time.Sleep(3 * time.Second)

	err = wb.Saved(true)
	if err != nil {
		t.Error(err)
	}

	err = wb.Close()
	if err != nil {
		t.Error(err)
	}
}
