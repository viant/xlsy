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
	names    []string
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
func (m *workSheet) ensureWorksheet() error {
	if m.index != nil {
		return nil
	}
	index, err := m.dest.NewSheet(m.name)
	if err != nil {
		return err
	}
	m.index = &index
	return nil
}

func (m *session) getOrCreateSheet(name string, first bool) (*workSheet, error) {
	if m.parent != nil {
		return m.parent.getOrCreateSheet(name, first)
	}
	m.mux.Lock()
	defer m.mux.Unlock()
	ret, ok := m.sheets[name]
	if ok {
		m.ensureFirst(name, first)
		return ret, nil
	}

	ret = &workSheet{name: name, dest: m.stylizer.file}
	m.sheets[name] = ret
	if first {
		m.names = append([]string{name}, m.names...)
	} else {
		m.names = append(m.names, name)
	}
	return ret, nil
}

func (m *session) ensureFirst(name string, first bool) {
	if first && m.names[0] != name {
		var names = []string{name}
		for _, item := range m.names {
			if item == name {
				continue
			}
			names = append(names, item)
		}
		m.names = names
	}
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
