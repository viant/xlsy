package xlsy

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strings"
)

// Stylizer represents stylizer
type Stylizer struct {
	defaultStyle string
	namedStyles  map[string]string
	file         *excelize.File
	registry     map[string]*Style
}

func (s *Stylizer) styleDefinition(def string, refs string) (string, error) {
	if s.defaultStyle == "" && def == "" && refs == "" {
		return "", nil
	}

	if refs != "" {
		for _, ref := range strings.Split(refs, " ") {
			aStyle, ok := s.namedStyles[ref]
			if !ok {
				return "", fmt.Errorf("failed to lookup ref style: %s", aStyle)
			}
			if def != "" {
				def += ";"
			}
			def += aStyle
		}
	}

	if s.defaultStyle == "" {
		return def, nil
	}
	if def == "" {
		return s.defaultStyle, nil
	}
	return s.defaultStyle + ";" + def, nil
}

// Style returns register style or nil
func (s *Stylizer) Style(style string) *Style {
	return s.registry[style]
}

// Register register a style
func (s *Stylizer) Register(style *Style) (err error) {
	prev, ok := s.registry[style.Definition]
	if ok {
		*style = *prev
		return
	}
	if style.ID != "" {
		prev, ok = s.registry[style.ID]
		if ok {
			*style = *prev
			return
		}
	}
	if err = style.Init(); err != nil {
		return err
	}
	s.registry[style.Definition] = style

	if style.ID != "" {
		s.registry[style.ID] = style
	}
	if err = s.register(style.Cell); err != nil {
		return err
	}

	err = s.register(style.Header)
	return err
}

func (s Stylizer) register(style *extStyle) error {
	if style.Style == nil {
		return nil
	}
	id, err := s.file.NewStyle(style.Style)
	if err != nil {
		return err
	}
	style.ID = &id
	return nil
}
