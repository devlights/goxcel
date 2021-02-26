# Goxcel

Goxcel is a library to operate Excel using [go-ole](https://github.com/go-ole/go-ole) package. Thanks [go-ole](https://github.com/go-ole/go-ole) package! 

This library works only on Windows.

[![CodeFactor](https://www.codefactor.io/repository/github/devlights/goxcel/badge)](https://www.codefactor.io/repository/github/devlights/goxcel)
![Goxcel - Go Version](https://img.shields.io/badge/go-1.16-blue.svg)
[![PkgGoDev](https://pkg.go.dev/badge/github.com/devlights/goxcel)](https://pkg.go.dev/github.com/devlights/goxcel)

## Install

```shell script
go get -u github.com/devlights/goxcel
```

## Usage

### Import statement

```go
import (
    "github.com/devlights/goxcel"
)
```

### HelloWorld

```go
func init() {
	log.SetFlags(log.Flags() &^ log.LstdFlags)
}

// main is entry point of this app.
//
// noinspection GoNilness
func main() {
	os.Exit(run())
}

func run() int {
	// 0. Initialize Goxcel
	quitGoxcelFn, err := goxcel.InitGoxcel()
	if err != nil {
		log.Println(err)
		return 1
	}

	defer quitGoxcelFn()

	// 1. Create new Goxcel instance.
	g, goxcelReleaseFn, err := goxcel.NewGoxcel()
	if err != nil {
		log.Println(err)
		return 2
	}

	// must call goxcel's release function when function exited
	// otherwise excel process was remained.
	defer goxcelReleaseFn()

	// optional settings
	visible := false
	_ = g.Silent(visible)

	// 2. Get Workbooks instance.
	wbs, err := g.Workbooks()
	if err != nil {
		log.Println(err)
		return 3
	}

	// 3. Add Workbook
	wb, wbReleaseFn, err := wbs.Add()
	if err != nil {
		log.Println(err)
		return 4
	}

	// call workbook's release funciton
	defer wbReleaseFn()

	// 4. Get Worksheet
	ws, err := wb.Sheets(1)
	if err != nil {
		log.Println(err)
		return 5
	}

	// 5. Get Cell
	c, err := ws.Cells(1, 1)
	if err != nil {
		log.Println(err)
		return 6
	}

	// 6. Set the value to cell
	err = c.SetValue("こんにちはWorld")
	if err != nil {
		log.Println(err)
		return 7
	}

	// optional. Display Excel and see the result.
	_ = g.SetVisible(true)
	time.Sleep(15 * time.Second)

	// Workbook::SetSaved(true) and Workbook::Close() is automatically called when `defer wbReleaseFn()`.
	// Excel::Quit() and Excel::Release() is automatically called when `defer goxcelReleaseFn()`.

	return 0
}
```

Also look at the examples directory :)

