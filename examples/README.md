# Examples

This directory contains following examples.

- copy_sheet
  - Example of Worksheet.CopySheet
- font_and_interior
  - Example of Cell.Font and Cell.Interior
- helloworld
  - Example of Helloworld
- pagebreaks
  - Example of Worksheet.HPageBreaks
- paste_picture
  - Example of Shapes.AddPicture
- printer_orientation_adjust
  - Example of Worksheet.PageSetup
- range_walk
  - Example of XlRange.Walk
- select first cell
  - Example of Worksheets.Item and Worksheet.Activate
- self_comobject_handling
  - Example of how to handling Com-Object with Goxcel
- shapes_numbering
  - Example of Worksheet.Shapes
- sheet_footer_adjust
  - Example of PageSetup.SetCenterFooter
- sheet_walk
  - Example of Worksheets.Walk
- sheet_zoom_adjust
  - Example of Window.SetZoom

## Run

Each examples can run with ```go run```.

- copy_sheet

```sh
$ go run github.com/devlights/goxcel/examples/copy_sheet -srcdir path/to/src/excel/dir 
```

After processing, a file named result.xlsx is generated in the current directory and contains all the sheets in the Excel file under the directory specified by the parameter.

- font_and_interior

```shell script
$ go run github.com/devlights/goxcel/examples/font_and_interior
```

- helloworld

```shell script
$ go run github.com/devlights/goxcel/examples/helloworld
```

- pagebreaks

```shell script
$ go run github.com/devlights/goxcel/examples/pagebreaks -f /path/to/excel-file
```

- printer_orientation_adjust

```shell script
$ go run github.com/devlights/goxcel/examples/printer_orientation_adjust -d /path/to/excel-dir -o [landscape|portrait]
```

- range_walk

```shell script
$ go run github.com/devlights/goxcel/examples/range_walk
```

- select_first_cell

```shell script
$ go run github.com/devlights/goxcel/examples/select_first_cell -d /path/to/excel-dir
```

- self_comobject_handling

```shell script
$ go run github.com/devlights/goxcel/examples/self_comobject_handling
```

- shapes_numbering

```shell script
$ go run github.com/devlights/goxcel/examples/shapes_numbering -f /path/to/excel-file
```

- sheet_footer_adjust

```shell script
$ go run github.com/devlights/goxcel/examples/sheet_footer_adjust -d /path/to/excel-dir -f footer-value -p sheet-name-pattern
```

- sheet_walk

```shell script
$ go run github.com/devlights/goxcel/sheet_walk
```

- sheet_zoom_adjust

```shell script
$ go run github.com/devlights/goxcel/sheet_zoom_adjust -d /path/to/excel-dir -z zoom-ratio
```
