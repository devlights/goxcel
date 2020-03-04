package main

import (
	"github.com/devlights/goxcel"
	"log"
	"os"
	"time"
)

func init() {
	log.SetFlags(log.Flags() &^ log.LstdFlags)
}

func main() {
	os.Exit(run())
}

func run() int {
	g, r, _ := goxcel.NewGoxcel()
	defer r()

	wbs, _ := g.Workbooks()
	wb, wbr, _ := wbs.Add()
	defer wbr()

	wss, _ := wb.WorkSheets()
	for i := 0; i < 10; i++ {
		_, _ = wss.AddLast()
	}

	errSheet, err := wss.Walk(func(ws *goxcel.Worksheet, index int) error {
		name, err := ws.Name()
		if err != nil {
			return err
		}

		log.Println(name)

		return nil
	})

	if err != nil {
		name := ""
		if errSheet != nil {
			name, _ = errSheet.Name()
		}

		log.Printf("%v [%s]\n", err, name)
		return 1
	}

	_ = g.SetVisible(true)
	time.Sleep(15 * time.Second)

	return 0
}
