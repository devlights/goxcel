package goxcel

import "testing"

func TestReleaser_Count(t *testing.T) {
	releaser := NewReleaser()

	f1 := func() error {
		return nil
	}

	f2 := func() error {
		return nil
	}

	releaser.Add(f1)
	releaser.Add(f2)

	if releaser.Count() != 2 {
		t.Errorf("want: %d\tgot: %d", 2, releaser.Count())
	}
}

func TestReleaser_Release(t *testing.T) {
	data1 := make([]string, 0, 0)
	data2 := []string{
		"f2",
		"f1",
	}

	releaser := NewReleaser()

	f1 := func() error {
		data1 = append(data1, "f1")
		return nil
	}

	f2 := func() error {
		data1 = append(data1, "f2")
		return nil
	}

	releaser.Add(f1)
	releaser.Add(f2)

	err := releaser.Release()
	if err != nil {
		t.Error(err)
	}

	for i := range data1 {
		if data1[i] != data2[i] {
			t.Errorf("want: %s\tgot: %s", data2[i], data1[i])
		}
	}
}
