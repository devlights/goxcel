package goxcel

import (
	"log"
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

func (r *Releaser) Count() int {
	return len(r.items)
}

func (r *Releaser) Release() error {

	// Pop elements
	for r.Count() > 0 {
		index := r.Count() - 1
		f := r.items[index]

		err := f()
		if err != nil {
			log.Println(err)
		}

		r.items[index] = nil
		r.items = r.items[:index]
	}

	r.items = make([]func() error, 0, 256)

	return nil
}
