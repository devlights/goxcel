GOCMD=go
GOBUILD=$(GOCMD) build
GOCLEAN=$(GOCMD) clean
GOTEST=$(GOCMD) test
GOGET=$(GOCMD) get
GORUN=$(GOCMD) run

PRJ_NAME=goxcel
GITHUB_USER=devlights
PKG_NAME=github.com/$(GITHUB_USER)/$(PRJ_NAME)

.PHONY: all
all: clean build test

.PHONY: build
build:
	$(GOBUILD) -v -race $(PKG_NAME)

.PHONY: test
test:
	$(GOTEST) -v  -p=1 ./...

.PHONY: clean
clean:
	$(GOCLEAN) $(CMD_PKG)

.PHONY: sheet_footer_adjust_example
sheet_footer_adjust_example:
	$(GORUN) examples/sheet_footer_adjust/sheet_footer_adjust.go -d ${TARGET_DIR} -f ${FOOTER}

.PHONY: printer_orientation_example
printer_orientation_example:
	$(GORUN) examples/printer_orientation_adjust/printer_orientation_adjust.go -d ${TARGET_DIR} -o ${ORIENTATION}

.PHONY: sheet_zoom_example
sheet_zoom_example:
	$(GORUN) examples/sheet_zoom_adjust/sheet_zoom_adjust.go -d ${TARGET_DIR} -z ${ZOOM}
