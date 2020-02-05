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

	defer r()

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
