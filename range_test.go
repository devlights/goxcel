package goxcel

import (
	"github.com/devlights/goxcel/testutil"
	"testing"
)

func TestXlRange_Count(t *testing.T) {
	testutil.Interval()
	defer testutil.Interval()

	g, r, err := NewGoxcel()
	if err != nil {
		t.Error(err)
	}

	defer r()

	wbs, _ := g.Workbooks()

	wb, wbReleaseFn, _ := wbs.Add()
	defer wbReleaseFn()

	ws, _ := wb.Sheets(1)

	xlRange, err := ws.Range(1, 1, 2, 2)
	if err != nil {
		t.Error(err)
	}

	count, err := xlRange.Count()
	if err != nil {
		t.Error(err)
	}

	i := int(count)
	if i != 4 {
		t.Errorf("want: 4\tgot: %d", i)
	}
}
