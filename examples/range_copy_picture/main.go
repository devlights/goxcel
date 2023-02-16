package main

import (
	"fmt"
	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
	"github.com/skanehira/clipboard-image/v2"
	"io"
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

	err := usedRange.CopyPicture(constants.XlScreen, constants.XlBitmap)
	if err != nil {
		return err
	}

	// Read image-binary from Clipboard
	// Thanks: https://github.com/skanehira/clipboard-image
	cr, err := clipboard.Read()
	if err != nil {
		return err
	}

	// Write to file
	file, err := os.Create("./image.png")
	if err != nil {
		return err
	}
	defer file.Close()

	if _, err := io.Copy(file, cr); err != nil {
		return err
	}

	return nil
}
