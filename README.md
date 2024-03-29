# Goxcel

Goxcel is a library to operate Excel using [go-ole](https://github.com/go-ole/go-ole) package. Thanks [go-ole](https://github.com/go-ole/go-ole) package! 

This library works only on Windows.

[![CodeFactor](https://www.codefactor.io/repository/github/devlights/goxcel/badge)](https://www.codefactor.io/repository/github/devlights/goxcel)
![Goxcel - Go Version](https://img.shields.io/badge/go-1.19-blue.svg)
[![PkgGoDev](https://pkg.go.dev/badge/github.com/devlights/goxcel)](https://pkg.go.dev/github.com/devlights/goxcel)

## Install

```sh
go get github.com/devlights/goxcel@latest
```

## Usage

### Import statement

```go
package main

import (
    "github.com/devlights/goxcel"
)

func init() {
	log.SetFlags(0)
}

// main is entry point of this app.
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
	quit := goxcel.MustInitGoxcel()
	defer quit()

	// 1. Create new Goxcel instance.
	excel, release := goxcel.MustNewGoxcel()

	// must call goxcel release function when function exited
	// otherwise excel process was remained.
	defer release()

	// optional settings
	visible := false
	excel.MustSilent(visible)

	// 2. Get Workbooks instance.
	wbs := excel.MustWorkbooks()

	// 3. Add Workbook
	wb, wbRelease := wbs.MustAdd()

	// call workbook's release function
	defer wbRelease()

	// 4. Get Worksheet
	ws := wb.MustSheets(1)

	// 5. Get Cell
	c := ws.MustCells(1, 1)

	// 6. Set the value to cell
	c.MustSetValue("こんにちはWorld")

	p := filepath.Join(os.TempDir(), "helloworld.xlsx")
	log.Printf("SAVE FILE: %s\n", p)

	// 7. Save
	wb.MustSaveAs(p)

	// Workbook::SetSaved(true) and Workbook::Close() is automatically called when `defer wbReleaseFn()`.
	// Excel::Quit() and Excel::Release() is automatically called when `defer release()`.

	return 0, p
}
```

Also look at the "examples" directory :)

