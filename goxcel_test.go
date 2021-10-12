package goxcel

import (
	"testing"
)

func TestGoxcel_MustMethods(t *testing.T) {
	f := MustInitGoxcel()
	defer f()

	g, r := MustNewGoxcel()
	defer r()

	g.MustSilent(true)

	wb := g.MustWorkbooks()

	t.Logf("goxcel: %v\tworkbooks: %v", g.ComObject(), wb.ComObject())
}

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
		name  string
		in    bool
		out   error
		value bool
	}{
		{"true", true, nil, true},
		{"false", false, nil, false},
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

			curval, err := g.DisplayAlerts()
			if err != nil {
				t.Errorf("DisplayAlerts return err (%v)", err)
			}

			if curval != c.value {
				t.Errorf("DisplayAlerts() return %v, want %v", curval, c.value)
			}
		})
	}
}

func TestGoxcel_SetEnableEvents(t *testing.T) {
	cases := []struct {
		name  string
		in    bool
		out   error
		value bool
	}{
		{"true", true, nil, true},
		{"false", false, nil, false},
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

			curval, err := g.EnableEvents()
			if err != nil {
				t.Errorf("EnableEvents return err (%v)", err)
			}

			if curval != c.value {
				t.Errorf("EnableEvents() return %v, want %v", curval, c.value)
			}
		})
	}
}

func TestGoxcel_SetScreenUpdating(t *testing.T) {
	cases := []struct {
		name  string
		in    bool
		out   error
		value bool
	}{
		{"true", true, nil, true},
		{"false", false, nil, false},
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

			curval, err := g.ScreenUpdating()
			if err != nil {
				t.Errorf("ScreenUpdating return err (%v)", err)
			}

			if curval != c.value {
				t.Errorf("ScreenUpdating() return %v, want %v", curval, c.value)
			}
		})
	}
}

func TestGoxcel_Silent(t *testing.T) {
	cases := []struct {
		name           string
		out            error
		displayAlerts  bool
		enableEvents   bool
		screenUpdating bool
	}{
		{"call method", nil, false, false, false},
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

			err = g.Silent(false)
			if err != c.out {
				t.Errorf("Silent() return %v", c.out)
			}

			enabled, err := g.DisplayAlerts()
			if err != nil {
				t.Errorf("DisplayAlerts() return err %v", err)
			}

			if enabled != c.displayAlerts {
				t.Errorf("DisplayAlerts() return %v, want %v", enabled, c.displayAlerts)
			}

			enabled, err = g.EnableEvents()
			if err != nil {
				t.Errorf("EnableEvents() return err %v", err)
			}

			if enabled != c.enableEvents {
				t.Errorf("EnableEvents() return %v, want %v", enabled, c.enableEvents)
			}

			enabled, err = g.ScreenUpdating()
			if err != nil {
				t.Errorf("ScreenUpdating() return err %v", err)
			}

			if enabled != c.screenUpdating {
				t.Errorf("ScreenUpdating() return %v, want %v", enabled, c.screenUpdating)
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

	_ = g.Silent(true)

	wb, err := g.Workbooks()
	if err != nil {
		t.Error(err)
	}

	if wb == nil {
		t.Errorf("wb is nil")
	}
}
