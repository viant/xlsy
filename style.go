package xlsy

import (
	"fmt"
	csscolorparser "github.com/mazznoer/csscolorparser"
	"github.com/viant/parsly"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

type (
	Length struct {
		Size float64
		Unit string
	}

	extStyle struct {
		ID *int
		*excelize.Style
		Width     *Length
		WidthMax  *Length
		Height    *Length
		HeightMax *Length
	}

	//Style represents a style
	Style struct {
		ID         string
		Definition string
		Header     *extStyle
		Cell       *extStyle
	}
)

var pixelRatio = 6.00
var ptRatio = 4.5

func (l *Length) Value() float64 {
	switch l.Unit {
	case "px":
		return l.Size / pixelRatio
	case "pt":
		return l.Size / ptRatio
	default:
		return l.Size
	}
}

func (e *extStyle) ensureStyle() {
	if e.Style != nil {
		return
	}
	e.Style = &excelize.Style{}
}

func (e *extStyle) ensureFont() {
	e.ensureStyle()
	if e.Font != nil {
		return
	}
	e.Font = &excelize.Font{}
}

func (e *extStyle) ensureAlignment() {
	e.ensureStyle()
	if e.Style.Alignment == nil {
		return
	}
	e.Style.Alignment = &excelize.Alignment{}
}

func (e *extStyle) ensureProtection() {
	e.ensureStyle()
	if e.Style.Protection == nil {
		return
	}
	e.Style.Protection = &excelize.Protection{}
}

// Init initializes style
func (s *Style) Init() error {
	s.Cell = &extStyle{}
	s.Header = &extStyle{}
	return ParseStyle(s.Definition, s)
}

func (s *Style) update(key, value string) (err error) {
	switch strings.ToLower(key) {
	case "background-color":
		return s.updateBackgroundColor(value, s.Cell)
	case "header-background-color":
		return s.updateBackgroundColor(value, s.Header)
	case "background-gradient":
		return s.updateBackgroundGradient(value, s.Cell)
	case "header-background-gradient":
		return s.updateBackgroundGradient(value, s.Header)
	case "font-style":
		s.updateFontStyle(value, s.Cell)
	case "header-font-style":
		s.updateFontStyle(value, s.Header)
	case "font-family":
		s.updateFontFamily(value, s.Cell)
	case "header-font-family":
		s.updateFontFamily(value, s.Header)

	case "color":
		return s.updateColor(value, s.Cell)
	case "header-color":
		return s.updateColor(value, s.Header)
	case "vertical-align":
		return s.updateVerticalAlign(value, s.Cell)
	case "header-vertical-align":
		return s.updateVerticalAlign(value, s.Header)
	case "text-align":
		return s.updateTextAlign(value, s.Cell)
	case "header-text-align":
		return s.updateTextAlign(value, s.Header)
	case "text-wrap":
		s.updateTextWrap(s.Cell)
	case "header-text-wrap":
		s.updateTextWrap(s.Header)
	case "text-indent":
		return s.updateTextIndent(value, s.Cell)
	case "header-text-indent":
		return s.updateTextIndent(value, s.Header)
	case "width":
		return s.updateWidth(value)
	case "width-max":
		return s.updateMaxWidth(value)
	case "height":
		return s.updateHeight(value, s.Cell)
	case "header-height":
		return s.updateHeight(value, s.Header)
	case "format":
		s.updateFormat(value, s.Cell)

	}
	return nil
}

func ensureColor(value string) (string, error) {
	color, err := csscolorparser.Parse(value)
	if err != nil {
		return "", err
	}
	return color.HexString(), nil
}

func (s *Style) updateTextWrap(style *extStyle) {
	style.ensureAlignment()
	style.Alignment.WrapText = true
}

func (s *Style) updateTextIndent(value string, style *extStyle) (err error) {
	style.ensureAlignment()
	style.Alignment.Indent, err = strconv.Atoi(value)
	if err != nil {
		return fmt.Errorf("invalid text-indent: %w, %v", err, value)
	}
	return err
}

func (s *Style) updateColor(value string, style *extStyle) (err error) {
	style.ensureFont()
	style.Font.Color, err = ensureColor(value)
	return err
}

func (s *Style) updateHeight(value string, style *extStyle) (err error) {
	if style.Height, err = parseLength(value); err != nil {
		return fmt.Errorf("invalid height: %w, %s", err, value)
	}
	return err
}

func (s *Style) updateWidth(value string) (err error) {
	if s.Cell.Width, err = parseLength(value); err != nil {
		return fmt.Errorf("invalid width: %w, %s", err, value)
	}
	s.Header.Width = s.Cell.Width
	return err
}

func (s *Style) updateMaxWidth(value string) (err error) {
	if s.Cell.WidthMax, err = parseLength(value); err != nil {
		return fmt.Errorf("invalid width-max: %w, %s", err, value)
	}
	s.Header.WidthMax = s.Cell.WidthMax
	return err
}

func (s *Style) updateVerticalAlign(value string, style *extStyle) error {
	value = strings.ToLower(value)
	style.ensureAlignment()
	if err := s.validateVerticalAlign(value); err != nil {
		return err
	}
	style.Alignment.Vertical = value
	return nil
}

func (s *Style) validateVerticalAlign(value string) error {
	switch value {
	case "top", "justify", "center", "distributed":
		return nil
	default:
		return fmt.Errorf("unsupported vertical-align: %s", value)
	}
}

func (s *Style) updateTextAlign(value string, style *extStyle) error {
	style.ensureAlignment()
	style.Alignment.Horizontal = value
	return nil
}

func (s *Style) validateTextAlign(value string) error {
	switch value {
	case "left", "right", "center",
		"fill", "justify", "centerContinuous", "distributed":
		return nil
	default:
		return fmt.Errorf("unsupported text-align: %s", value)
	}
}

func (s *Style) updateBackgroundColor(value string, style *extStyle) (err error) {
	style.ensureStyle()
	style.Fill.Type = "solid"
	style.Fill.Color, err = ensureColors(value)
	return err
}

func ensureColors(value string) ([]string, error) {
	var colors []string
	for _, candidate := range strings.Split(value, " ") {
		color, err := ensureColor(candidate)
		if err != nil {
			return nil, err
		}
		colors = append(colors, color)
	}
	return colors, nil
}

func (s *Style) updateBackgroundGradient(value string, style *extStyle) (err error) {
	style.ensureStyle()
	style.Fill.Type = "gradient"
	style.Fill.Color, err = ensureColors(value)
	return err
}

func (s *Style) updateFontStyle(value string, style *extStyle) {
	style.ensureFont()
	for _, item := range strings.Split(value, " ") {
		switch strings.ToLower(item) {
		case "bold":
			style.Font.Bold = true
		case "italic":
			style.Font.Italic = true
		case "strike":
			style.Font.Strike = true
		}
	}
}

var (
	formatUsd     = "$#,##0.00"
	formatGbp     = "£#,##0.00"
	formatEur     = "€#,##0.00"
	formatCad     = "C$#,##0.00"
	formatAud     = "A$#,##0.00"
	formatDate    = "m/d/yy HH:mm:ss"
	formatISO8601 = "yyyy/mm/dd HH:mm:ss"
	formatPct     = "0%"
)

func (s *Style) updateFormat(value string, style *extStyle) {
	style.ensureStyle()
	switch strings.ToLower(value) {
	case "usd":
		style.Style.CustomNumFmt = &formatUsd
	case "gbp":
		style.Style.CustomNumFmt = &formatGbp
	case "eur":
		style.Style.CustomNumFmt = &formatEur
	case "cad":
		style.Style.CustomNumFmt = &formatCad
	case "aud":
		style.Style.CustomNumFmt = &formatAud
	case "date":
		style.Style.CustomNumFmt = &formatDate
	case "iso8601":
		style.Style.CustomNumFmt = &formatISO8601
	case "pct":
		style.Style.CustomNumFmt = &formatPct
	default:
		style.Style.CustomNumFmt = &value
	}
}

func (s *Style) updateFontFamily(value string, style *extStyle) {
	style.ensureFont()
	style.Font.Family = value
}

func parseLength(length string) (*Length, error) {
	cursor := parsly.NewCursor("", []byte(length), 0)
	match := cursor.MatchAny(numberMatcher)
	if match.Code != numberToken {
		return nil, cursor.NewError(numberMatcher)
	}
	ret := &Length{Unit: "px"}
	var err error
	ret.Size, err = strconv.ParseFloat(match.Text(cursor), 32)
	if err != nil {
		return nil, err
	}
	if cursor.Pos < len(cursor.Input) {
		unit := string(cursor.Input[cursor.Pos:])
		switch strings.ToLower(unit) {
		case "px":
			ret.Unit = "px"
		case "pt":
			ret.Unit = "pt"
		case "":
		default:
			return nil, fmt.Errorf("unsupported Length unit: '%s'", unit)
		}
	}
	return ret, nil
}

func ParseStyle(definition string, style *Style) error {
	if definition == "" {
		return nil
	}

	cursor := parsly.NewCursor("", []byte(definition), 0)
	for cursor.Pos < len(cursor.Input) {
		match := cursor.MatchOne(colonTerminatorMatcher)
		if match.Code == parsly.EOF {
			return nil
		}
		if match.Code != colonTerminatorToken {
			return cursor.NewError(colonTerminatorMatcher)
		}
		key, value := matchStylePair(cursor, match)
		if key == "" {
			return nil
		}
		if err := style.update(key, value); err != nil {
			return err
		}
	}
	return nil
}

func matchStylePair(cursor *parsly.Cursor, match *parsly.TokenMatch) (string, string) {

	key := match.Text(cursor)
	key = key[:len(key)-1] //exclude =
	value := ""
	match = cursor.MatchAny(singleQuotedMatcher, semicolonTerminatorMatcher)
	switch match.Code {
	case singleQuotedToken:
		value = match.Text(cursor)
		value = value[1 : len(value)-1]
		match = cursor.MatchAny(comaTerminatorMatcher)
	case semicolonTerminatorToken:
		value = match.Text(cursor)
		value = value[:len(value)-1] //exclude ,

	default:
		if cursor.Pos < len(cursor.Input) {
			value = string(cursor.Input[cursor.Pos:])
		}
		cursor.Pos = len(cursor.Input)
	}
	return key, value
}
