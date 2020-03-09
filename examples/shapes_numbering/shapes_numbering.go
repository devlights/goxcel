package main

import (
	"flag"
	"fmt"
	"github.com/devlights/goxcel"
	"log"
	"os"
	"sort"
)

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

		ss, _ := ws.Shapes()

		shapeCount, _ := ss.Count()
		shapes := make([]*goxcel.Shape, 0, shapeCount)
		for i := 1; i <= int(shapeCount); i++ {
			s, _ := ss.Item(i)
			shapes = append(shapes, s)
		}

		sort.Slice(shapes, func(i, j int) bool {
			lRange, _ := shapes[i].TopLeftCell()
			rRange, _ := shapes[j].TopLeftCell()

			lRow, _ := lRange.Row()
			rRow, _ := rRange.Row()

			return lRow < rRow
		})

		for i, s := range shapes {
			topLeft, _ := s.TopLeftCell()

			c, _ := topLeft.Column()
			r, _ := topLeft.Row()

			col := int(c)
			row := int(r - 1)
			if row <= 0 {
				row = 1
			}

			cell, _ := ws.Cells(row, col)
			_ = cell.SetValue(fmt.Sprintf("No.%02d", i+1))
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
