package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/goxcel"
)

var (
	targetDirectory string
)

func main() {
	os.Exit(run())
}

func run() int {
	flag.StringVar(&targetDirectory, "d", "", "対象ディレクトリ (必須)")
	flag.Parse()

	if targetDirectory == "" {
		flag.Usage()
		return 2
	}

	err := filepath.Walk(targetDirectory, walkFiles)
	if err != nil {
		log.Println(err)
		return 1
	}

	return 0
}

func walkFiles(path string, info os.FileInfo, err error) error {
	if err != nil {
		return err
	}

	if info.IsDir() {
		return nil
	}

	if !strings.HasSuffix(path, "xlsx") {
		return nil
	}

	absPath, err := filepath.Abs(path)
	if err != nil {
		return err
	}

	quitGoxcelFn, _ := goxcel.InitGoxcel()
	defer quitGoxcelFn()

	g, r, err := goxcel.NewGoxcel()
	if err != nil {
		return err
	}

	defer r()

	_ = g.Silent(false)

	wbs, err := g.Workbooks()
	if err != nil {
		return err
	}

	wb, wbReleaseFn, err := wbs.Open(absPath)
	if err != nil {
		return err
	}
	defer wbReleaseFn()

	wss, err := wb.WorkSheets()
	if err != nil {
		return err
	}

	errorWs, err := wss.Walk(walkSheets)
	if err != nil {
		errSheetName := ""
		if errorWs != nil {
			errSheetName, _ = errorWs.Name()
		}

		err = fmt.Errorf("%w at sheet[%s]", err, errSheetName)
		return err
	}

	ws, err := wss.Item(1)
	if err != nil {
		return err
	}

	err = ws.Activate()
	if err != nil {
		return err
	}

	err = wb.Save()
	if err != nil {
		return err
	}

	return nil
}

//noinspection GoUnusedParameter
func walkSheets(ws *goxcel.Worksheet, index int) error {
	err := ws.Activate()
	if err != nil {
		return err
	}

	c, err := ws.Cells(1, 1)
	if err != nil {
		return err
	}

	err = c.Select()
	if err != nil {
		return err
	}

	return nil
}
