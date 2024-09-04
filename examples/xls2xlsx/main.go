package main

import (
	"flag"
	"io/fs"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
)

type (
	Args struct {
		dir        string
		rmOriginal bool
		verbose    bool
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
	flag.BoolVar(&args.rmOriginal, "rm", false, "remove original file")
	flag.BoolVar(&args.verbose, "verbose", false, "verbose mode")
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

		if !strings.HasSuffix(path, ".xls") {
			return nil
		}

		wb, wbRelease, err := wbs.Open(abs(path))
		if err != nil {
			return err
		}
		defer wbRelease()

		if args.verbose {
			appLog.Printf("converting: %s", path)
		}

		err = wb.SaveAsWithFileFormat(rename(path), constants.XlOpenXMLWorkbook)
		if err != nil {
			return err
		}

		if args.rmOriginal {
			return os.Remove(path)
		}

		return nil
	})

	if err != nil {
		return err
	}

	return nil
}

func rename(p string) string {
	var (
		base = filepath.Base(p)
		dir  = filepath.Dir(p)
		xlsx = abs(filepath.Join(dir, strings.ReplaceAll(base, "xls", "xlsx")))
	)

	return xlsx
}
