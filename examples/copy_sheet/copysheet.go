package main

import (
	"flag"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/goxcel"
)

// flag parameters
var (
	srcDir string
	out    string
)

// logs
var (
	wsLog = log.New(os.Stdout, ">>> ", 0)
)

func main() {
	os.Exit(run())
}

func run() int {
	flag.StringVar(&srcDir, "srcdir", "", "source directory")
	flag.StringVar(&out, "out", "result.xlsx", "output file name")
	flag.Parse()

	if srcDir == "" {
		flag.Usage()
		return 1
	}

	g, r, _ := goxcel.NewGoxcel()
	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(false)

	wbs, _ := g.Workbooks()
	wbDest, wbDestR, _ := wbs.Add()
	defer wbDestR()

	wsDest, _ := wbDest.Sheets(1)
	_ = filepath.Walk(srcDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if info.IsDir() {
			return nil
		}

		absPath, _ := filepath.Abs(path)
		if !strings.HasSuffix(absPath, "xlsx") {
			return nil
		}

		wsLog.Println(path)

		wb, wbr, _ := wbs.Open(absPath)
		defer wbr()

		wss, _ := wb.WorkSheets()
		_, err = wss.Walk(func(ws *goxcel.Worksheet, index int) error {
			err := ws.CopySheet(wsDest, false)
			if err != nil {
				return err
			}

			return nil
		})

		return err
	})

	wd, _ := os.Getwd()
	curdir, _ := filepath.Abs(wd)
	resultPath := filepath.Join(curdir, out)

	err := wbDest.SaveAs(resultPath)
	if err != nil {
		log.Println(err)
		return 2
	}

	return 0
}
