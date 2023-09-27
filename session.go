package xlsy

import (
	"sync"
)

// Option represents marshaller session
type Option func(m *session) error

type session struct {
	parent   *session
	tag      *Tag
	stylizer *Stylizer
	sheets   map[string]*workSheet
	mux      sync.Mutex
}

func (m *session) apply(options []Option) error {
	for _, opt := range options {
		if err := opt(m); err != nil {
			return err
		}
	}
	if m.tag == nil {
		m.tag = &Tag{}
	}
	return nil
}

func (m *session) getOrCreateSheet(name string) (*workSheet, error) {
	if m.parent != nil {
		return m.parent.getOrCreateSheet(name)
	}
	m.mux.Lock()
	defer m.mux.Unlock()
	ret, ok := m.sheets[name]
	if ok {
		return ret, nil
	}
	index, err := m.stylizer.file.NewSheet(name)
	if err != nil {
		return nil, err
	}
	ret = &workSheet{name: name, index: index, dest: m.stylizer.file}
	m.sheets[name] = ret
	return ret, nil
}

func newSession(parent *session, stylizer *Stylizer, tag *Tag) *session {
	return &session{
		parent:   parent,
		stylizer: stylizer,
		sheets:   map[string]*workSheet{},
		mux:      sync.Mutex{},
		tag:      tag,
	}
}

// WithTag return tag session
func WithTag(tag *Tag) Option {
	return func(m *session) error {
		m.tag = tag
		return nil
	}
}

// WithInverted return vertical option
func WithInverted() Option {
	t := true
	return func(m *session) error {
		m.tag.Inverted = &t
		return nil
	}
}

// WithNamedStyles accept name/style definition pairs
func WithNamedStyles(pairs ...string) Option {
	return func(m *session) error {
		if m.stylizer.namedStyles == nil {
			m.stylizer.namedStyles = make(map[string]string)
		}
		for i := 0; i < len(pairs); i += 2 {
			m.stylizer.namedStyles[pairs[0]] = pairs[1]
		}
		return nil
	}
}

func WithDefaultHeaderStyle(definition string) Option {
	return func(m *session) error {
		m.stylizer.defaultHeaderStyle = definition
		return nil
	}
}

func WithDefaultStyle(definition string) Option {
	return func(m *session) error {
		m.stylizer.defaultCellStyle = definition
		return nil
	}
}
