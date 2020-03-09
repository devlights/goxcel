package main

import (
	"flag"
	"fmt"
	"github.com/devlights/goxcel"
	"log"
	"os"
)

func init() {
	log.SetFlags(log.Flags() &^ log.LstdFlags)
}

func main() {
	os.Exit(run())
}

func run() int {
	var (
		targetFilePath string
	)

	flag.StringVar(&targetFilePath, "f", "", "対象Excelファイルの絶対パス (必須)")
	flag.Parse()

	if targetFilePath == "" {
		flag.Usage()
		return 1
	}

	if _, err := os.Stat(targetFilePath); err != nil {
		log.Println(err)
		return 2
	}

	quitGoxcelFn, _ := goxcel.InitGoxcel()
	defer quitGoxcelFn()

	g, r, _ := goxcel.NewGoxcel()
	defer r()

	wbs, _ := g.Workbooks()
	wb, wbr, err := wbs.Open(targetFilePath)
	if err != nil {
		log.Println(err)
		return 3
	}
	defer wbr()

	wss, _ := wb.WorkSheets()
	_, err = wss.Walk(func(ws *goxcel.Worksheet, index int) error {

		hpbs, _ := ws.HPageBreaks()
		_, err = hpbs.Walk(func(hpb *goxcel.HPageBreak, index int) error {
			location, err := hpb.Location()
			if err != nil {
				return err
			}

			value := fmt.Sprintf("Page.%d", index+1)
			err = location.SetValue(value)
			if err != nil {
				return err
			}

			return nil
		})

		if err != nil {
			return err
		}

		return nil
	})

	if err != nil {
		log.Println(err)
		return 4
	}

	_ = wb.Save()

	return 0
}
