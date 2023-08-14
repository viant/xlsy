package xlsy

import (
	"fmt"
	"github.com/viant/parsly"
	"strconv"
	"strings"
)

const TagName = "xls"

type Tag struct {
	Name      string
	Style     string
	StyleRef  string
	Ignore    bool
	Blank     bool
	Position  *int
	Embed     bool
	Direction string
	OffsetX   int
	OffsetY   int
}

func (t *Tag) update(key, value string) error {
	switch strings.ToLower(key) {
	case "name":
		t.Name = value
	case "embed":
		t.Embed = true
	case "direction", "dir":
		t.Direction = strings.ToLower(value)
	case "offsetx":
		if err := convertAndSetInt(&t.OffsetX, "offsetX", value); err != nil {
			return err
		}
	case "offsety":
		if err := convertAndSetInt(&t.OffsetY, "offsetY", value); err != nil {
			return err
		}
	case "pos":
		var pos int
		if err := convertAndSetInt(&pos, "pos", value); err != nil {
			return err
		}
		t.Position = &pos
	case "style":
		t.Style = value
	case "styleref":
		t.StyleRef = value
	case "":
		return nil
	default:
		return fmt.Errorf("unsupported xls sheet tag key: %s", key)
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
	case "blank":
		ret.Blank = true
		return ret, nil
	case "embed":
		ret.Embed = true
	}
	cursor := parsly.NewCursor("", []byte(tag), 0)
	for cursor.Pos < len(cursor.Input) {
		key, value := matchPair(cursor)
		if err := ret.update(key, value); err != nil {
			return nil, err
		}
		if key == "" {
			break
		}
	}
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
