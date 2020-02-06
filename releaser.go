package goxcel

import (
	"log"
	"sort"
)

type (
	Releaser struct {
		items []func() error
	}
)

func NewReleaser() *Releaser {
	r := new(Releaser)
	r.items = make([]func() error, 0, 256)

	return r
}

func (r *Releaser) Add(f func() error) {
	r.items = append(r.items, f)
}

func (r *Releaser) Release() error {
	// reverse sort
	sort.Slice(r.items, func(i, j int) bool {
		return !(i < j)
	})

	for _, f := range r.items {
		err := f()
		if err != nil {
			log.Println(err)
		}
	}

	r.items = make([]func() error, 0, 256)

	return nil
}
