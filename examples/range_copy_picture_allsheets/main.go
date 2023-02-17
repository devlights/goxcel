package main

import (
	"bufio"
	"fmt"
	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
	"os"
	"path/filepath"
)

func main() {
	dir, err := os.Getwd()
	if err != nil {
		panic(err)
	}

	if err := run(dir); err != nil {
		panic(err)
	}
}

//goland:noinspection GoUnhandledErrorResult
func run(curdir string) error {
	quit := goxcel.MustInitGoxcel()
	defer quit()

	excel, release := goxcel.MustNewGoxcel()
	defer release()

	wbs := excel.MustWorkbooks()
	wb, wbRelease := wbs.MustOpen(filepath.Join(curdir, "testdata", "TestData.xlsx"))
	defer wbRelease()

	wss := wb.MustWorkSheets()
	_, err := wss.Walk(func(ws *goxcel.Worksheet, index int) error {
		name, err := ws.Name()
		if err != nil {
			return err
		}

		usedRange, err := ws.UsedRange()
		if err != nil {
			return err
		}

		file, err := os.Create(filepath.Join(curdir, fmt.Sprintf("%s.png", name)))
		if err != nil {
			return err
		}
		defer file.Close()

		bufW := bufio.NewWriter(file)
		err = usedRange.CopyPictureToFile(bufW, constants.XlScreen, constants.XlBitmap)
		if err != nil {
			return err
		}
		defer bufW.Flush()

		return nil
	})

	if err != nil {
		return err
	}

	return nil
}
