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

const (
	OrientationPortrait  = "portrait"
	OrientationLandscape = "landscape"
)

var (
	targetDirectory string
	sheetPattern    string
	orientation     string
)

func main() {
	flag.StringVar(&targetDirectory, "d", "", "対象ディレクトリ (必須)")
	flag.StringVar(&sheetPattern, "p", "", "シート名の条件、指定しない場合は全シートが対象")
	flag.StringVar(&orientation, "o", OrientationPortrait, "印刷方向 (portrait(縦) or landscape(横)) (必須)")
	flag.Parse()

	if targetDirectory == "" || orientation == "" {
		flag.Usage()
		os.Exit(2)
	}

	if orientation != OrientationPortrait && orientation != OrientationLandscape {
		flag.Usage()
		os.Exit(3)
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

	if sheetPattern != "" {
		name, _ := ws.Name()
		if !strings.Contains(name, sheetPattern) {
			return nil
		}
	}

	ps, err := ws.PageSetup()
	if err != nil {
		return err
	}

	o := goxcel.XlPageOrientationXlPortrait
	if orientation == OrientationLandscape {
		o = goxcel.XlPageOrientationXlLandscape
	}

	err = ps.SetOrientation(o)
	if err != nil {
		return err
	}

	return nil
}
