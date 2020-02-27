# Goxcel

Goxcel is a library to operate Excel using [go-ole](https://github.com/go-ole/go-ole) package. Thanks [go-ole](https://github.com/go-ole/go-ole) package! 

This library works only on Windows.

[![CodeFactor](https://www.codefactor.io/repository/github/devlights/goxcel/badge)](https://www.codefactor.io/repository/github/devlights/goxcel)
![Goxcel - Go Version](https://img.shields.io/badge/go-1.13-blue.svg)

## Install

```shell script
go get -u github.com/devlights/goxcel
```

## Usage

### Import statement

```go
import "github.com/devlights/goxcel"
```

### HelloWorld

```go
    // 1. Create new Goxcel instance.
    g, goxcelReleaseFn, err := goxcel.NewGoxcel()
    if err != nil {
        // error process
    }
    
    // must call goxcel's release function when function exited
    // otherwise excel process was remained.
    defer goxcelReleaseFn()
    
    // optional settings
    _ = g.SetDisplayAlerts(false)
    _ = g.SetVisible(true)
    
    // 2. Get Workbooks instance.
    wbs, err := g.Workbooks()
    if err != nil {
        // error process
    }
    
    // 3. Add Workbook
    wb, wbReleaseFn, err := wbs.Add()
    if err != nil {
        // error process
    }
    
    // call workbook's release funciton
    defer wbReleaseFn()
    
    // 4. Get Worksheet
    ws, err := wb.Sheets(1)
    if err != nil {
        // error process
    }
    
    // 5. Get Cell
    c, err := ws.Cells(1, 1)
    if err != nil {
        // error process
    }
    
    // 6. Set the value to cell
    err = c.SetValue("helloworld")
    if err != nil {
        // error process
    }
    
    // 7. Call the Workbook::SetSaved method to not show a dialog on exit
    err = wb.SetSaved(true)
    if err != nil {
        // error process
    }
    
    // 8. Close Workbook
    err = wb.Close()
    if err != nil {
        // error process
    }
```

Also look at the examples directory :)

