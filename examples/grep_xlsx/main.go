package main

import (
	"flag"
	"io/fs"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/goxcel"
)

type (
	Args struct {
		dir     string
		text    string
		onlyHit bool
		verbose bool
		debug   bool
	}
)

var (
	args Args
)

var (
	appLog = log.New(os.Stdout, "", 0)
)

func init() {
	flag.StringVar(&args.dir, "dir", ".", "directory path")
	flag.StringVar(&args.text, "text", "", "search text")
	flag.BoolVar(&args.onlyHit, "only-hit", true, "show ONLY HIT")
	flag.BoolVar(&args.verbose, "verbose", false, "verbose mode")
	flag.BoolVar(&args.debug, "debug", false, "debug mode")
}

func abs(p string) string {
	abs, err := filepath.Abs(p)
	if err != nil {
		log.Panic(err)
	}

	return abs
}

func main() {
	log.SetFlags(0)
	flag.Parse()

	if args.text == "" {
		flag.PrintDefaults()
		os.Exit(1)
	}

	if args.dir == "" {
		args.dir = "."
	}

	if err := run(); err != nil {
		log.Fatal(err)
	}
}

func run() error {
	quit := goxcel.MustInitGoxcel()
	defer quit()

	excel, release := goxcel.MustNewGoxcel()
	defer release()

	excel.MustSilent(false)

	wbs := excel.MustWorkbooks()

	rootDir := abs(args.dir)
	err := filepath.WalkDir(rootDir, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}

		if d.IsDir() {
			return nil
		}

		if strings.Contains(filepath.Base(path), "~$") {
			return nil
		}

		if !strings.HasSuffix(path, ".xlsx") {
			return nil
		}

		absPath := abs(path)
		wb, wsRelease, err := wbs.Open(absPath)
		if err != nil {
			return err
		}
		defer wsRelease()

		if args.debug {
			appLog.Printf("Document Open: %s", absPath)
		}

		sheets, err := wb.WorkSheets()
		if err != nil {
			return err
		}

		relPath, _ := filepath.Rel(rootDir, absPath)
		_, err = sheets.Walk(func(ws *goxcel.Worksheet, index int) error {
			rng, err := ws.UsedRange()
			if err != nil {
				return err
			}

			startCell, err := rng.Cells(1, 1)
			if err != nil {
				return err
			}

			foundRange, found, err := rng.Find(args.text, startCell)
			if err != nil {
				return err
			}

			if !found {
				if !args.onlyHit {
					appLog.Printf("%s: Not Found", relPath)
				}

				return nil
			}

			name, _ := ws.Name()
			if args.verbose {
				col, _ := foundRange.Column()
				row, _ := foundRange.Row()
				value, _ := foundRange.Value()

				appLog.Printf("%s %q (%d,%d): %q", relPath, name, row, col, value)
			} else {
				appLog.Printf("%s: %q", relPath, name)
				return nil
			}

			startCell, _ = foundRange.Cells(1, 1)
			for i := 0; found; i++ {
				after, _ := foundRange.Cells(1, 1)

				foundRange, found, err = rng.FindNext(after)
				if err != nil {
					return err
				}

				if !found {
					break
				}

				if args.debug {
					printCellPos(startCell, after)
				}

				if i > 0 && sameCell(startCell, after) {
					break
				}

				if args.verbose {
					col, _ := foundRange.Column()
					row, _ := foundRange.Row()
					value, _ := foundRange.Value()

					appLog.Printf("%s %q (%d,%d): %q", relPath, name, row, col, value)
				}
			}

			return nil
		})

		if err != nil {
			return err
		}

		return nil
	})

	if err != nil {
		return err
	}

	return nil
}

func sameCell(c1, c2 *goxcel.Cell) bool {
	col1, _ := c1.Column()
	row1, _ := c1.Row()
	col2, _ := c2.Column()
	row2, _ := c2.Row()

	return col1 == col2 && row1 == row2
}

func printCellPos(startCell, after *goxcel.Cell) {
	col1, _ := startCell.Column()
	row1, _ := startCell.Row()
	col2, _ := after.Column()
	row2, _ := after.Row()

	appLog.Printf("FindNext: startCell=(%d,%d)\tafter=(%d,%d)", row1, col1, row2, col2)
}
