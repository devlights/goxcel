package main

import (
	"github.com/devlights/goxcel"
	"log"
	"time"
)

func init() {
	log.SetFlags(log.Flags() &^ log.LstdFlags)
}

// main is entry point of this app.
//
// noinspection GoNilness
func main() {
	// 1. Create new Goxcel instance.
	g, goxcelReleaseFn, err := goxcel.NewGoxcel()
	if err != nil {
		log.Println(err)
		return
	}

	// must call goxcel's release function when function exited
	// otherwise excel process was remained.
	defer goxcelReleaseFn()

	// optional settings
	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(false)

	// 2. Get Workbooks instance.
	wbs, err := g.Workbooks()
	if err != nil {
		log.Println(err)
		return
	}

	// 3. Add Workbook
	wb, wbReleaseFn, err := wbs.Add()
	if err != nil {
		log.Println(err)
		return
	}

	// call workbook's release funciton
	defer wbReleaseFn()

	// 4. Get Worksheet
	ws, err := wb.Sheets(1)
	if err != nil {
		log.Println(err)
		return
	}

	// 5. Get Cell
	c, err := ws.Cells(1, 1)
	if err != nil {
		log.Println(err)
		return
	}

	// 6. Set the value to cell
	err = c.SetValue("こんにちはWorld")
	if err != nil {
		log.Println(err)
		return
	}

	// optional. Display Excel and see the result.
	_ = g.SetVisible(true)
	time.Sleep(15 * time.Second)

	// 7. Call the Workbook::SetSaved method to not show a dialog on exit
	err = wb.SetSaved(true)
	if err != nil {
		log.Println(err)
		return
	}

	// 8. Close Workbook
	err = wb.Close()
	if err != nil {
		log.Println(err)
		return
	}
}
