package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/devlights/goxcel"
)

var (
	appLog = log.New(os.Stdout, "", 0)
	errLog = log.New(os.Stderr, "[ERROR] ", 0)
)

func main() {
	if err := run(); err != nil {
		errLog.Panic(err)
	}
}

func run() error {
	quit := goxcel.MustInitGoxcel()
	defer quit()

	excel, release := goxcel.MustNewGoxcel()
	defer release()

	excel.MustSetVisible(true)

	wbs := excel.MustWorkbooks()
	wb, wbRelease := wbs.MustAdd()
	defer wbRelease()

	ws := wb.MustSheets(1)
	for row := 1; row <= 100; row++ {
		for col := 1; col <= 100; col++ {
			c := ws.MustCells(row, col)
			c.MustSetValue(fmt.Sprintf("%v_%v", row, col))
		}
	}

	// 正しく最大行数が取得できているかを確認するために
	// 明示的に5行ほど空行を空けてから、再度行追加

	for row := 105; row <= 200; row++ {
		for col := 1; col <= 100; col++ {
			c := ws.MustCells(row, col)
			c.MustSetValue(fmt.Sprintf("%v_%v", row, col))
		}
	}

	maxRow, maxCol, err := ws.MaxRowCol(1, 1)
	if err != nil {
		return err
	}
	fmt.Printf("maxRow=%v\tmaxCol=%v\n", maxRow, maxCol)

	<-time.After(5 * time.Second)

	return nil
}
