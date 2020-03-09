package main

import (
	"log"
	"os"
	"time"

	"github.com/devlights/goxcel"
	"golang.org/x/sync/errgroup"
)

func init() {
	log.SetFlags(log.Flags() &^ log.LstdFlags)
}

func main() {
	os.Exit(run())
}

func run() int {
	var (
		errGrp = errgroup.Group{}
	)

	// Goxel works on STA mode.
	// runtime.LockOSThread() is called inside goxcel.InitGoxcel().
	errGrp.Go(func() error {
		// 0. Initialize Goxcel
		//    Lock current goroutine thread for STA behavior.
		quitGoxcelFn, err := goxcel.InitGoxcel()
		if err != nil {
			return err
		}

		// Unlock thread lock
		defer quitGoxcelFn()

		// 1. Create new Goxcel instance.
		g, goxcelReleaseFn, err := goxcel.NewGoxcel()
		if err != nil {
			return err
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
			return err
		}

		// 3. Add Workbook
		wb, wbReleaseFn, err := wbs.Add()
		if err != nil {
			return err
		}

		// call workbook's release funciton
		defer wbReleaseFn()

		// 4. Get Worksheet
		ws, err := wb.Sheets(1)
		if err != nil {
			return err
		}

		// 5. Get Cell
		c, err := ws.Cells(1, 1)
		if err != nil {
			return err
		}

		// 6. Set the value to cell
		err = c.SetValue("こんにちはWorld")
		if err != nil {
			return err
		}

		// optional. Display Excel and see the result.
		_ = g.SetVisible(true)
		time.Sleep(15 * time.Second)

		// 7. Call the Workbook::SetSaved method to not show a dialog on exit
		err = wb.SetSaved(true)
		if err != nil {
			return err
		}

		// 8. Close Workbook
		err = wb.Close()
		if err != nil {
			return err
		}

		return nil
	})

	if err := errGrp.Wait(); err != nil {
		log.Println(err)
		return 1
	}

	return 0
}
