package goxcel

import (
	"testing"
)

func TestHyperLinks_Add_Item_Count_Delete(t *testing.T) {
	//
	// Arrange
	//
	quit := MustInitGoxcel()
	defer quit()

	excel, release := MustNewGoxcel()
	defer release()

	wbs := excel.MustWorkbooks()
	wb, wbr := wbs.MustAdd()
	defer wbr()

	ws := wb.MustSheets(1)

	hl, err := ws.HyperLinks()
	if err != nil {
		t.Error(err)
	}

	ra, err := ws.Range(1, 1, 1, 1)
	if err != nil {
		t.Error(err)
	}

	//
	// Act
	//
	err = hl.Add(ra, "https://www.google.co.jp", "", "screenTip", "textToDisplay")
	if err != nil {
		t.Error(err)
	}

	//
	// Assert
	//
	hl2, err := hl.Item(1)
	if err != nil {
		t.Error(err)
	}

	if hl2 == nil {
		t.Errorf("[want] not nil\t[got] nil")
	}

	count, err := hl.Count()
	if err != nil {
		t.Error(err)
	}

	if count != 1 {
		t.Errorf("[want] 1\t[got] %v", count)
	}

	//
	// Act2
	//
	err = hl.Delete()
	if err != nil {
		t.Error(err)
	}

	//
	// Assert2
	//
	count, err = hl.Count()
	if err != nil {
		t.Error(err)
	}

	if count != 0 {
		t.Errorf("[want] 0\t[got] %v", count)
	}
}
