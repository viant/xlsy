package xlsy

import (
	"fmt"
	"github.com/viant/parsly"
	"strconv"
	"strings"
)

const TagName = "xls"

type (
	StyleTag struct {
		Destination string
		Style       string
		Ref         string
	}

	Tag struct {
		Name         string
		WorkSheet    string
		HeaderStyle  *StyleTag
		CellStyle    *StyleTag
		ColumnStyle  *StyleTag
		Embed        bool
		Ignore       bool
		Blank        bool
		Omitempty    bool
		Position     *int
		Inverted     *bool //inverted orientation
		Row          int
		Column       int
		ColumnOffset int
		RowOffset    int
	}
)

func (t *Tag) Invert() bool {
	if t.Inverted == nil {
		return false
	}
	return *t.Inverted
}

func (s *StyleTag) Definition() *Style {
	if s == nil {
		return nil
	}

	return &Style{Definition: s.Style, Destination: s.Destination}
}

func (t *Tag) updateStyles(styles []*StyleTag) {
	if len(styles) == 0 {
		return
	}
	for _, candidate := range styles {
		destStyle := t.ensureDestination(candidate)
		setStyle(candidate, destStyle)

	}
}

func (t *Tag) adjustAddress(addr *Cursor) {
	if t.Row != 0 {
		addr.setRow(t.Row)
	}
	if t.Column != 0 {
		addr.setColumn(t.Column)
	}
	if t.RowOffset != 0 {
		addr.incRow(t.RowOffset)
	}
	if t.ColumnOffset != 0 {
		addr.incColumn(t.ColumnOffset)
	}
}

func (t *Tag) ensureDestination(candidate *StyleTag) *StyleTag {
	var destStyle *StyleTag
	switch candidate.Destination {
	case "header":
		if t.HeaderStyle == nil {
			t.HeaderStyle = &StyleTag{}
		}
		destStyle = t.HeaderStyle
	case "column":
		if t.ColumnStyle == nil {
			t.ColumnStyle = &StyleTag{}
		}
		destStyle = t.ColumnStyle
	default:
		if t.CellStyle == nil {
			t.CellStyle = &StyleTag{}
		}
		destStyle = t.CellStyle
	}
	destStyle.Destination = candidate.Destination
	return destStyle
}

func setStyle(candidate, t *StyleTag) {
	if candidate.Style != "" {
		if t.Style != "" {
			t.Style += ";"
		}
		t.Style = candidate.Style
	}
	if candidate.Style != "" {
		if t.Ref != "" {
			t.Ref += ","
		}
		t.Ref = candidate.Ref
	}
}

func (t *Tag) update(key, value string, styles *[]*StyleTag) error {
	attr := strings.ToLower(key)
	switch attr {
	case "name":
		t.Name = value
	case "worksheet":
		t.WorkSheet = value
	case "omitempty":
		t.Omitempty = true
	case "embed":
		embed := value == "true"
		t.Embed = embed
	case "invert", "inverted":
		invert := value == "true"
		t.Inverted = &invert
	case "row":
		if err := convertAndSetInt(&t.Row, "row", value); err != nil {
			return err
		}
	case "column":
		if err := convertAndSetInt(&t.Column, "column", value); err != nil {
			return err
		}
	case "columoffset":
		if err := convertAndSetInt(&t.ColumnOffset, "columOffset", value); err != nil {
			return err
		}
	case "rowoffset":
		if err := convertAndSetInt(&t.RowOffset, "rowOffset", value); err != nil {
			return err
		}
	case "pos":
		var pos int
		if err := convertAndSetInt(&pos, "pos", value); err != nil {
			return err
		}
		t.Position = &pos
	default:
		if strings.HasSuffix(attr, "style") {
			var style = &StyleTag{Style: value}
			if index := strings.Index(attr, "."); index != -1 {
				style.Destination = attr[:index]
			}
			*styles = append(*styles, style)
		} else if strings.HasSuffix(attr, "styleref") {
			var style = &StyleTag{Ref: value}
			if index := strings.Index(attr, "."); index != -1 {
				style.Destination = attr[:index]
			}
			*styles = append(*styles, style)
		} else {
			return fmt.Errorf("unsupported xls tag key: %s", key)
		}
	case "":
		return nil
	}
	return nil
}

func parseTag(tag string) (*Tag, error) {
	ret := &Tag{}
	if tag == "" {
		return ret, nil
	}
	switch strings.ToLower(tag) {
	case "-":
		ret.Ignore = true
		return ret, nil
	case ",blank":
		ret.Blank = true
		return ret, nil
	case ",embed":
		ret.Embed = true
	case ",omitempty":
		ret.Omitempty = true
	}

	var styles []*StyleTag
	cursor := parsly.NewCursor("", []byte(tag), 0)
	for cursor.Pos < len(cursor.Input) {
		key, value := matchPair(cursor)
		if err := ret.update(key, value, &styles); err != nil {
			return nil, err
		}
		if key == "" {
			break
		}
	}
	ret.updateStyles(styles)
	return ret, nil
}

func matchPair(cursor *parsly.Cursor) (string, string) {
	match := cursor.MatchOne(equalTerminatorMatcher)
	if match.Code != equalTerminatorToken {
		return "", ""
	}
	key := match.Text(cursor)
	key = key[:len(key)-1] //exclude =
	value := ""
	match = cursor.MatchAny(scopeBlockMatcher, comaTerminatorMatcher)
	switch match.Code {
	case scopeBlockToken:
		value = match.Text(cursor)
		value = value[1 : len(value)-1]
		match = cursor.MatchAny(comaTerminatorMatcher)
	case comaTerminatorToken:
		value = match.Text(cursor)
		value = value[:len(value)-1] //exclude ,

	default:
		if cursor.Pos < len(cursor.Input) {
			value = string(cursor.Input[cursor.Pos:])
			cursor.Pos = len(cursor.Input)
		}
	}

	return key, value
}

func convertAndSetInt(dest *int, field, value string) error {
	intValue, err := strconv.Atoi(value)
	if err != nil {
		return fmt.Errorf("invalud %s: %s, %w", field, value, err)
	}
	*dest = intValue
	return nil
}
