package main

import (
	"fmt"
	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
	"os"
)

func main() {
	if err := run(); err != nil {
		panic(err)
	}
}

//goland:noinspection GoUnhandledErrorResult
func run() error {
	quit := goxcel.MustInitGoxcel()
	defer quit()

	excel, release := goxcel.MustNewGoxcel()
	defer release()

	wbs := excel.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	rows := []int{1, 2, 3, 4, 5, 10, 20, 30}
	for _, row := range rows {
		cols := []int{1, 2, 3, 4, 5}
		for _, col := range cols {
			c, _ := ws.Cells(row, col)
			c.MustSetValue(fmt.Sprintf("%v,%v", row, col))
		}
	}

	usedRange, _ := ws.UsedRange()
	_ = usedRange.Select()

	file, err := os.Create("./image.png")
	if err != nil {
		return err
	}
	defer file.Close()

	err = usedRange.CopyPictureToFile(file, constants.XlScreen, constants.XlBitmap)
	if err != nil {
		return err
	}

	return nil
}
