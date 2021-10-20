# Goxcel

Goxcel is a library to operate Excel using [go-ole](https://github.com/go-ole/go-ole) package. Thanks [go-ole](https://github.com/go-ole/go-ole) package! 

This library works only on Windows.

[![CodeFactor](https://www.codefactor.io/repository/github/devlights/goxcel/badge)](https://www.codefactor.io/repository/github/devlights/goxcel)
![Goxcel - Go Version](https://img.shields.io/badge/go-1.17-blue.svg)
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
	ret, xlsx := run()
	if ret == 0 {
		// Launch EXCEL
		_ = exec.Command("cmd", "/C", xlsx).Run()
	}

	os.Exit(ret)
}

func run() (int, string) {
	// 0. Initialize Goxcel
	quitGoxcelFn := goxcel.MustInitGoxcel()
	defer quitGoxcelFn()

	// 1. Create new Goxcel instance.
	g, goxcelReleaseFn := goxcel.MustNewGoxcel()

	// must call goxcel's release function when function exited
	// otherwise excel process was remained.
	defer goxcelReleaseFn()

	// optional settings
	visible := false
	g.MustSilent(visible)

	// 2. Get Workbooks instance.
	wbs := g.MustWorkbooks()

	// 3. Add Workbook
	wb, wbReleaseFn := wbs.MustAdd()

	// call workbook's release funciton
	defer wbReleaseFn()

	// 4. Get Worksheet
	ws := wb.MustSheets(1)

	// 5. Get Cell
	c := ws.MustCells(1, 1)

	// 6. Set the value to cell
	if err := c.SetValue("こんにちはWorld"); err != nil {
		log.Println(err)
		return 6, ""
	}

	p := filepath.Join(os.TempDir(), "helloworld.xlsx")
	log.Printf("SAVE FILE: %s\n", p)

	// 7. Save
	if err := wb.SaveAs(p); err != nil {
		log.Println(err)
		return 7, ""
	}

	// Workbook::SetSaved(true) and Workbook::Close() is automatically called when `defer wbReleaseFn()`.
	// Excel::Quit() and Excel::Release() is automatically called when `defer goxcelReleaseFn()`.

	return 0, p
}
```

Also look at the examples directory :)

