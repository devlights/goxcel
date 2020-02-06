package goxcel

import (
	"github.com/devlights/goxcel/testutil"
	"testing"
)

func TestGoxcel_Startup(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	err = g.SetDisplayAlerts(false)
	if err != nil {
		t.Error(err)
	}

	err = g.SetVisible(true)
	if err != nil {
		t.Error(err)
	}
}

func TestGoxcel_Workbooks(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()

	if err != nil {
		t.Error(err)
	}

	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(true)

	wb, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	if wb == nil {
		t.Errorf("wb is nil")
	}
}
