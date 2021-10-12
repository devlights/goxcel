package main

import (
	"log"
	"os"
	"os/exec"
	"path/filepath"

	"github.com/devlights/goxcel"
)

func init() {
	log.SetFlags(log.Flags() &^ log.LstdFlags)
}

// main is entry point of this app.
//
// noinspection GoNilness
func main() {
	ret, xlsx := run()
	if ret == 0 {
		// Launch EXCEL
		_ = exec.Command("cmd", "/C", xlsx).Run()
	}

	os.Exit(ret)
}

func run() (int, string) {
	// 0. Initialize Goxcel
	quitGoxcelFn := goxcel.MustInitGoxcel()
	defer quitGoxcelFn()

	// 1. Create new Goxcel instance.
	g, goxcelReleaseFn := goxcel.MustNewGoxcel()

	// must call goxcel's release function when function exited
	// otherwise excel process was remained.
	defer goxcelReleaseFn()

	// optional settings
	visible := false
	g.MustSilent(visible)

	// 2. Get Workbooks instance.
	wbs := g.MustWorkbooks()

	// 3. Add Workbook
	wb, wbReleaseFn := wbs.MustAdd()

	// call workbook's release funciton
	defer wbReleaseFn()

	// 4. Get Worksheet
	ws := wb.MustSheets(1)

	// 5. Get Cell
	c := ws.MustCells(1, 1)

	// 6. Set the value to cell
	if err := c.SetValue("こんにちはWorld"); err != nil {
		log.Println(err)
		return 6, ""
	}

	p := filepath.Join(os.TempDir(), "goxcel_must_methods.xlsx")
	log.Printf("SAVE FILE: %s\n", p)

	// 7. Save
	if err := wb.SaveAs(p); err != nil {
		log.Println(err)
		return 7, ""
	}

	// Workbook::SetSaved(true) and Workbook::Close() is automatically called when `defer wbReleaseFn()`.
	// Excel::Quit() and Excel::Release() is automatically called when `defer goxcelReleaseFn()`.

	return 0, p
}
