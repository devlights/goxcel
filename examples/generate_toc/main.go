package main

import (
	"fmt"
	"github.com/devlights/goxcel"
	"path/filepath"
)

func main() {
	if err := run(); err != nil {
		panic(err)
	}
}

func run() error {
	//
	// Initialize Goxcel
	//
	quit := goxcel.MustInitGoxcel()
	defer quit()

	excel, release := goxcel.MustNewGoxcel()
	defer release()

	wbs := excel.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	//
	// Add 20 worksheet
	//
	const WSCOUNT = 20
	wss := wb.MustWorkSheets()
	wsList := make([]*goxcel.Worksheet, 0, 10)

	for i := 0; i < WSCOUNT; i++ {
		ws, _ := wss.AddLast()
		name, _ := ws.Name()

		cell := ws.MustCells(1, 1)
		cell.MustSetValue(name)

		wsList = append(wsList, ws)
	}

	//
	// Add TOC sheet
	//
	firstSheet, _ := wss.Item(1)
	tocSheet, _ := wss.AddBefore(firstSheet)
	_ = tocSheet.SetName("TOC")

	tocCell := tocSheet.MustCells(1, 1)
	tocCell.MustSetValue("Table Of Contents")

	//
	// Make TOC
	//
	for i, ws := range wsList {
		row := i + 2

		hl, _ := tocSheet.HyperLinks()
		ra, _ := tocSheet.Range(row, 1, row, 1)

		wsName, _ := ws.Name()
		addr := ""
		subAddr := fmt.Sprintf("'%s'!A1", wsName)
		screenTip := ""

		_ = hl.Add(ra, addr, subAddr, screenTip, wsName)
	}

	//
	// Write Workbook
	//
	abs, _ := filepath.Abs(".")
	fpath := filepath.Join(abs, "result.xlsx")

	_ = wb.SaveAs(fpath)

	return nil
}
