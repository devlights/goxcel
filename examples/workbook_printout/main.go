package main

import (
	"fmt"
	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
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

	wss := wb.MustWorkSheets()
	ws, _ = wss.AddLast()
	for row := 1; row <= 10; row++ {
		for col := 1; col <= 5; col++ {
			c := ws.MustCells(row, col)
			c.MustSetValue(fmt.Sprintf("hello-%v_%v", row, col))
		}
	}

	cnt, _ := wss.Count()
	appLog.Printf("sheet count=%v\n", cnt)

	filePath := "c:/tmp/aaa.pdf"
	if _, err := os.Stat(filePath); !os.IsNotExist(err) {
		_ = os.Remove(filePath)
	}

	if err := wb.ExportAsFixedFormat(constants.XlTypePDF, filePath); err != nil {
		return err
	}

	if err := wb.PrintOut(); err != nil {
		return err
	}

	return nil
}
