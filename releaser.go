package goxcel

import "log"

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
	for _, f := range r.items {
		err := f()
		if err != nil {
			log.Println(err)
		}
	}

	return nil
}
