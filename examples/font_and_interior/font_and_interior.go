package main

import (
	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
	"log"
	"time"
)

func init() {
	log.SetFlags(log.Flags() &^ log.LstdFlags)
}

// main is entry point of this app.
//
// noinspection GoNilness
func main() {
	g, r, _ := goxcel.NewGoxcel()
	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(false)

	wbs, _ := g.Workbooks()
	wb, wbr, _ := wbs.Add()
	defer wbr()

	ws, _ := wb.Sheets(1)
	cell, _ := ws.Cells(1, 1)
	_ = cell.SetValue("こんにちはWorld")

	font, _ := cell.Font()
	_ = font.SetColor(constants.RgbRed)
	_ = font.SetBold(true)
	_ = font.SetItalic(true)
	_ = font.SetUnderline(constants.XlUnderlineStyleSingle)
	_ = font.SetSize(40)
	_ = font.SetName("ＭＳ ゴシック")

	interior, _ := cell.Interior()
	_ = interior.SetColor(constants.RgbBlue)

	// optional. Display Excel and see the result.
	_ = g.SetVisible(true)
	time.Sleep(15 * time.Second)
}
