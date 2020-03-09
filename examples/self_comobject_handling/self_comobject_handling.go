package main

import (
	"fmt"
	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/util"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"log"
	"os"
)

func main() {
	os.Exit(run())
}

func run() int {
	quitGoxcelFn, _ := goxcel.InitGoxcel()
	defer quitGoxcelFn()

	g, r, _ := goxcel.NewGoxcel()
	defer r()

	wbs, _ := g.Workbooks()

	wb, wbr, _ := wbs.Add()
	defer wbr()

	ws, _ := wb.Sheets(1)

	value, _, err := util.ProcessComObject(ws, func(c *ole.IDispatch) (interface{}, *ole.IDispatch, error) {
		v, err := oleutil.GetProperty(c, "Name")
		if err != nil {
			return nil, nil, err
		}

		return v.Value(), v.ToIDispatch(), nil
	})

	if err != nil {
		log.Println(err)
		return 1
	}

	fmt.Println(value) // --> Sheet1

	return 0
}
