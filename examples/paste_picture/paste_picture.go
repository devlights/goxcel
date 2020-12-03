package main

import (
	"flag"
	"image"
	_ "image/png"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/goxcel"
)

var (
	srcdir  string
	workdir string
	outfile string
	debug   bool
)

var (
	appLog, errLog, dbgLog *log.Logger
)

func init() {
	wd, _ := os.Getwd()
	ab, _ := filepath.Abs(wd)
	workdir = ab

	flag.StringVar(&srcdir, "srcdir", "", "The directory where the image files are located (Required)")
	flag.StringVar(&outfile, "out", filepath.Join(workdir, "result.xlsx"), "output file name")
	flag.BoolVar(&debug, "debug", false, "enable debug mode")
}

func main() {
	flag.Parse()

	appLog, errLog, dbgLog = log.New(os.Stdout, "", 0), log.New(os.Stderr, "", 0), log.New(os.Stdout, "", 0)
	if !debug {
		dbgLog.SetOutput(ioutil.Discard)
	}

	if srcdir == "" {
		errLog.Println("xlsimgpaste.exe -srcdir path-to-img-files-directory")
		flag.PrintDefaults()
		return
	}

	srcdir, _ = filepath.Abs(srcdir)
	outfile, _ = filepath.Abs(outfile)

	os.Exit(run())
}

func run() int {
	dbgLog.Println(workdir)

	f, _ := goxcel.InitGoxcel()
	defer f()

	g, r, _ := goxcel.NewGoxcel()
	defer r()

	_ = g.Silent(false)

	wbs, _ := g.Workbooks()

	wb, wbr, _ := wbs.Add()
	defer wbr()

	ws, _ := wb.Sheets(1)
	shapes, _ := ws.Shapes()

	left, top := 10, 10
	err := filepath.Walk(srcdir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if info.IsDir() {
			return nil
		}

		if !strings.HasSuffix(path, "png") {
			return nil
		}

		dbgLog.Println(path)
		width, height, _ := getImgSize(path)

		err = shapes.AddPicture(path, left, top, width, height)
		if err != nil {
			errLog.Println(err)
			return err
		}

		top += height + 10

		return nil
	})

	if err != nil {
		errLog.Println(err)
		return 1
	}

	_ = wb.SaveAs(outfile)

	return 0
}

func getImgSize(filepath string) (width, height int, err error) {
	f, err := os.Open(filepath)
	if err != nil {
		return 0, 0, err
	}

	im, _, err := image.DecodeConfig(f)
	if err != nil {
		return 0, 0, err
	}

	return im.Width, im.Height, nil
}
