package main

import (
	"flag"
	"fmt"
	"github.com/devlights/goxcel"
	"log"
	"os"
	"path/filepath"
	"strings"
)

var (
	targetDirectory string
)

func main() {
	flag.StringVar(&targetDirectory, "d", "", "対象ディレクトリ (必須)")
	flag.Parse()

	if targetDirectory == "" {
		flag.Usage()
		os.Exit(2)
	}

	err := filepath.Walk(targetDirectory, walkFiles)
	if err != nil {
		log.Fatal(err)
	}

	os.Exit(0)
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

	g, r, err := goxcel.NewGoxcel()
	if err != nil {
		return err
	}

	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)

	wbs, err := g.Workbooks()
	if err != nil {
		return err
	}

	wb, err := wbs.Open(absPath)
	if err != nil {
		return err
	}

	defer func() {
		_ = wb.SetSaved(true)
		_ = wb.Close()
	}()

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

func walkSheets(ws *goxcel.Worksheet) error {
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
