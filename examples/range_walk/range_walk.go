package main

import (
	"fmt"
	"github.com/devlights/goxcel"
	"log"
	"time"
)

func init() {
	log.SetFlags(log.Flags() &^ log.LstdFlags)
}

func main() {
	g, r, _ := goxcel.NewGoxcel()
	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)

	wbs, _ := g.Workbooks()

	wb, wbReleaseFn, _ := wbs.Add()
	defer wbReleaseFn()

	ws, _ := wb.Sheets(1)

	myRange, _ := ws.Range(1, 1, 2, 2)
	_, err := myRange.Walk(func(r *goxcel.XlRange, c *goxcel.Cell, row, col int) error {
		e := c.SetValue(fmt.Sprintf("'%d:%d", row, col))
		if e != nil {
			return e
		}

		return nil
	})

	if err != nil {
		log.Fatal(err)
	}

	time.Sleep(10 * time.Second)

	_, err = myRange.Walk(func(r *goxcel.XlRange, c *goxcel.Cell, row, col int) error {
		v, e := c.Value()
		if e != nil {
			return e
		}

		log.Printf("row: %d\tcol: %d\tvalue: %v\n", row, col, v)
		return nil
	})

	if err != nil {
		log.Fatal(err)
	}
}
