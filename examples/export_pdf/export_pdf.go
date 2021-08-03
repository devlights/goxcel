package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
)

// flag parameters
var (
	src string
	dst string
)

// logs
var (
	appLog = log.New(os.Stdout, ">>> ", 0)
)

func main() {
	var (
		returnCode int
	)

	if err := run(); err != nil {
		_, _ = fmt.Fprint(os.Stderr, err)
		returnCode = -1
	}

	appLog.Println("done")

	os.Exit(returnCode)
}

func run() error {
	abs := func(p string) string {
		v, _ := filepath.Abs(p)
		return v
	}

	flag.StringVar(&src, "src", "", "source file")
	flag.StringVar(&dst, "dst", "result.pdf", "output pdf name")
	flag.Parse()

	if src == "" {
		flag.Usage()
		return nil
	}

	quitFn, _ := goxcel.InitGoxcel()
	defer quitFn()

	g, r, _ := goxcel.NewGoxcel()
	defer r()

	_ = g.Silent(false)

	wbs, err := g.Workbooks()
	if err != nil {
		return err
	}

	absPath := abs(src)
	wb, wbReleaseFn, err := wbs.Open(abs(src))
	if err != nil {
		return err
	}
	defer wbReleaseFn()
	appLog.Printf("WorkBook Open: %s", absPath)

	ws, err := wb.Sheets(1)
	if err != nil {
		return err
	}

	absPath = abs(src)
	err = ws.ExportAsFixedFormat(constants.XlTypePDF, absPath)
	if err != nil {
		return err
	}
	appLog.Printf("Export   PDF : %s", absPath)

	return nil
}
