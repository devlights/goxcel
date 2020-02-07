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

.PHONY: footer_adjust_example
footer_adjust_example:
	$(GORUN) examples/sheet_footer_adjust/footer_adjust.go -d ${TARGET_DIR} -p ${SHEET_PATTERN} -f ${FOOTER}
