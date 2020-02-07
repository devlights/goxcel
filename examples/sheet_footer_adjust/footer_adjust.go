package main

import (
	"flag"
	"github.com/devlights/goxcel"
	"log"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	var (
		targetDirectory string
		sheetPattern    string
		footer          string
	)

	flag.StringVar(&targetDirectory, "d", "", "対象ディレクトリ (必須)")
	flag.StringVar(&sheetPattern, "p", "", "シート名の条件、指定しない場合は全シートが対象")
	flag.StringVar(&footer, "f", "&P", "フッターに設定する値")
	flag.Parse()

	if targetDirectory == "" {
		flag.Usage()
		os.Exit(2)
	}

	err := filepath.Walk(targetDirectory, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if info.IsDir() {
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

		sheetCount, _ := wss.Count()
		for i := 0; i < int(sheetCount); i++ {
			ws, err := wss.Item(i + 1)
			if err != nil {
				return err
			}

			err = ws.Activate()
			if err != nil {
				return err
			}

			if sheetPattern != "" {
				name, _ := ws.Name()
				if !strings.Contains(name, sheetPattern) {
					continue
				}
			}

			ps, err := ws.PageSetup()
			if err != nil {
				return err
			}

			err = ps.SetCenterFooter(footer)
			if err != nil {
				return err
			}
		}

		err = wb.Save()
		if err != nil {
			return err
		}

		return nil
	})

	if err != nil {
		log.Fatal(err)
	}

	os.Exit(0)
}
