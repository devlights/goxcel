package goxcel

import (
	"testing"
	"time"
)

func TestGoxcelStartup(t *testing.T) {
	a := NewArgs("")
	g, r, err := NewGoxcel(a)

	if err != nil {
		t.Error(err)
	}

	defer r(false)

	err = g.Visible(true)
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	err = g.Quit()
	if err != nil {
		t.Error(err)
	}
}

func TestGoxcelWorkbooks(t *testing.T) {
	a := NewArgs("")
	g, r, err := NewGoxcel(a)

	if err != nil {
		t.Error(err)
	}

	defer r(true)

	_ = g.Visible(true)

	wb, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	b, err := wb.Add()
	if err != nil {
		t.Error(err)
	}

	time.Sleep(3 * time.Second)

	err = b.Saved(true)
	if err != nil {
		t.Error(err)
	}

	err = b.Close()
	if err != nil {
		t.Error(err)
	}
}
