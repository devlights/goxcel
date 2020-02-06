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

	err = g.SetDisplayAlerts(false)
	if err != nil {
		t.Error(err)
	}

	err = g.SetVisible(true)
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)
}

func TestGoxcelWorkbooksAdd(t *testing.T) {
	time.Sleep(3 * time.Second)

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

	b, err := wb.Add()
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	err = b.SetSaved(true)
	if err != nil {
		t.Error(err)
	}

	err = b.Close()
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)
}

func TestGoxcelWorkbooksOpen(t *testing.T) {
	time.Sleep(3 * time.Second)

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
	b, err := wb.Open(xlsxPath)
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	err = b.SetSaved(true)
	if err != nil {
		t.Error(err)
	}

	err = b.Close()
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)
}

func TestGoxcelWorkbookSave(t *testing.T) {
	time.Sleep(3 * time.Second)

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
	wb, _ := wbs.Open(xlsxPath)

	ws, _ := wb.Sheets(1)
	c, _ := ws.Cells(1, 1)
	_ = c.SetValue("こんにちわ goxcel")

	err = wb.Save()
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	_ = wb.SetSaved(true)
	_ = wb.Close()

	time.Sleep(3 * time.Second)
}

func TestGoxcelWorkbookSaveAs(t *testing.T) {
	time.Sleep(3 * time.Second)

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
	wb, _ := wbs.Open(srcXlsxPath)

	ws, _ := wb.Sheets(1)
	c, _ := ws.Cells(1, 1)
	_ = c.SetValue("hello goxcel")

	dstXlsxPath := filepath.Join(userHomeDir, "Book2.xlsx")
	err = wb.SaveAs(dstXlsxPath)
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	_ = wb.SetSaved(true)
	_ = wb.Close()

	time.Sleep(3 * time.Second)
}

func TestGoxcelCellValue(t *testing.T) {
	time.Sleep(3 * time.Second)

	g, r, err := NewGoxcel()
	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)
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

	err = wb.SetSaved(true)
	if err != nil {
		t.Error(err)
	}

	err = wb.Close()
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)
}
