package main

import (
	"fmt"
	"github.com/devlights/goxcel"
	"log"
	"os"
)

var (
	appLog = log.New(os.Stdout, "", 0)
	errLog = log.New(os.Stderr, "[ERROR] ", 0)
)

func main() {
	if err := run(); err != nil {
		errLog.Panic(err)
	}
}

func run() error {
	quit := goxcel.MustInitGoxcel()
	defer quit()

	excel, release := goxcel.MustNewGoxcel()
	defer release()

	excel.MustSetVisible(true)

	wbs := excel.MustWorkbooks()
	wb, wbRelease := wbs.MustAdd()
	defer wbRelease()

	ws := wb.MustSheets(1)
	for row := 1; row <= 10; row++ {
		for col := 1; col <= 5; col++ {
			c := ws.MustCells(row, col)
			c.MustSetValue(fmt.Sprintf("%v_%v", row, col))
		}
	}

	if err := ws.PrintOut(); err != nil {
		return err
	}

	appLog.Println("DONE")

	return nil
}
