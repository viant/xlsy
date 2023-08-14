package xlsy

// Option represents marshaller options
type Option func(m *options) error

type options struct {
	tag      *Tag
	stylizer *Stylizer
}

func (m *options) apply(options []Option) error {
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

// WithTag return tag options
func WithTag(tag *Tag) Option {
	return func(m *options) error {
		m.tag = tag
		return nil
	}
}

// WithNamedStyles accept name/style definition pairs
func WithNamedStyles(pairs ...string) Option {
	return func(m *options) error {
		if m.stylizer.namedStyles == nil {
			m.stylizer.namedStyles = make(map[string]string)
		}
		for i := 0; i < len(pairs); i += 2 {
			m.stylizer.namedStyles[pairs[0]] = pairs[1]
		}
		return nil
	}
}

func WithDefaultStyle(definition string) Option {
	return func(m *options) error {
		m.stylizer.defaultStyle = definition
		return nil
	}
}
