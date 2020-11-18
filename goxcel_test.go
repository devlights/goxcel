package goxcel

import (
	"testing"
)

func TestGoxcel_Startup(t *testing.T) {
	f, err := InitGoxcel()
	if err != nil {
		t.Fatal(err)
	}
	defer f()

	g, r, err := NewGoxcel()
	if err != nil {
		t.Fatal(err)
	}
	defer r()

	t.Logf("goxcel: %v", g.ComObject())
}

func TestGoxcel_SetDisplayAlerts(t *testing.T) {
	cases := []struct {
		name string
		in   bool
		out  error
	}{
		{"true", true, nil},
		{"false", false, nil},
	}

	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			f, err := InitGoxcel()
			if err != nil {
				t.Fatal(err)
			}
			defer f()

			g, r, err := NewGoxcel()
			if err != nil {
				t.Fatal(err)
			}
			defer r()

			err = g.SetDisplayAlerts(c.in)
			if err != c.out {
				t.Errorf("SetDisplayAlerts(%v) return err (%v)", c.in, err)
			}
		})
	}
}

func TestGoxcel_SetEnableEvents(t *testing.T) {
	cases := []struct {
		name string
		in   bool
		out  error
	}{
		{"true", true, nil},
		{"false", false, nil},
	}

	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			f, err := InitGoxcel()
			if err != nil {
				t.Fatal(err)
			}
			defer f()

			g, r, err := NewGoxcel()
			if err != nil {
				t.Fatal(err)
			}
			defer r()

			err = g.SetEnableEvents(c.in)
			if err != c.out {
				t.Errorf("SetEnableEvents(%v) return err (%v)", c.in, err)
			}
		})
	}
}

func TestGoxcel_SetScreenUpdating(t *testing.T) {
	cases := []struct {
		name string
		in   bool
		out  error
	}{
		{"true", true, nil},
		{"false", false, nil},
	}

	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			f, err := InitGoxcel()
			if err != nil {
				t.Fatal(err)
			}
			defer f()

			g, r, err := NewGoxcel()
			if err != nil {
				t.Fatal(err)
			}
			defer r()

			err = g.SetScreenUpdating(c.in)
			if err != c.out {
				t.Errorf("SetScreenUpdating(%v) return err (%v)", c.in, err)
			}
		})
	}
}

func TestGoxcel_Workbooks(t *testing.T) {
	f, err := InitGoxcel()
	if err != nil {
		t.Fatal(err)
	}
	defer f()

	g, r, err := NewGoxcel()
	if err != nil {
		t.Fatal(err)
	}
	defer r()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetEnableEvents(false)
	_ = g.SetScreenUpdating(false)
	_ = g.SetVisible(true)

	wb, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	if wb == nil {
		t.Errorf("wb is nil")
	}
}
