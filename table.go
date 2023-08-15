package xlsy

import (
	"github.com/viant/xunsafe"
	"reflect"
	"sort"
	"strconv"
)

const (

	//DirectionHorizontal = "horizontal"
	//DirectionVertical vertial direction
	DirectionVertical = "vertical"
)

const (
	defaultSheetName = "Sheet1"
)

type (
	//Table represents a table
	Table struct {
		Parent   *Column
		Stylizer *Stylizer
		Tag      *Tag
		Columns  Columns
		Type     reflect.Type
	}
	//ColumnAddress represents acolumn
	Column struct {
		Position int
		Name     string
		Tag      *Tag
		Field    *xunsafe.Field
		xType    *xunsafe.Type
		Table    *Table
		size     int
	}

	//Columns represents columns
	Columns []*Column
)

// Size represents a size
func (c *Column) Size() int {
	if c.size != 0 {
		return c.size
	}
	if c.Table == nil {
		c.size = 1
		return c.size
	}
	c.size = c.Table.ColumnSize()
	return c.size
}

// HeaderStyleID returns header style ID or nil
func (c *Column) HeaderStyleID(stylizer *Stylizer) *int {
	if c.Tag.Style == "" {
		return nil
	}
	if c.Tag.Style != "" {
		if style := stylizer.Style(c.Tag.Style); style != nil && style.Header.Style != nil {
			return style.Header.ID
		}
	}
	return nil
}

// Width returns width
func (c *Column) Width(stylizer *Stylizer) *Length {
	if c.Tag.Style == "" {
		return nil
	}
	if c.Tag.Style != "" {
		if style := stylizer.Style(c.Tag.Style); style != nil && style.Cell != nil {
			width := style.Cell.Width
			if style.Cell.WidthMax != nil {
				if style.Cell.WidthMax.Value() < width.Value() {
					return style.Cell.WidthMax
				}
			}
			return style.Cell.Width
		}
	}
	return nil
}

// CellStyleID returns a style ID
func (c *Column) CellStyleID(stylizer *Stylizer) *int {
	if c.Tag.Style == "" {
		return nil
	}
	if c.Tag.Style != "" {
		if style := stylizer.Style(c.Tag.Style); style != nil && style.Cell.Style != nil {
			return style.Cell.ID
		}
	}
	return nil
}

// Style returns column style
func (c *Column) Style() *Style {
	if c.Tag.Style == "" {
		return nil
	}
	return &Style{Definition: c.Tag.Style}
}

// ColumnSize returns column size
func (t *Table) ColumnSize() int {
	ret := 0
	for _, column := range t.Columns {
		ret += column.Size()
	}
	return ret
}

// IsVertical returns true if table direction is vertical
func (t *Table) IsVertical() bool {
	if t.Tag.Direction == DirectionVertical {
		return true
	}
	return false
}

// SheetName returns table cheed name
func (t *Table) SheetName() string {
	if t.Tag.Name != "" {
		return t.Tag.Name
	}
	return defaultSheetName
}

// Address returns cell address
func (t *Table) Address(row, column int) string {
	address := t.address(row, column)
	return address
}

// ColumnAddress returns column address
func (t *Table) ColumnAddress(column int) string {
	return string([]byte{'A' + byte(column)})
}

func (t *Table) address(row int, column int) string {
	if t.IsVertical() {
		y := strconv.Itoa(1 + column)
		x := 'A' + byte(row)
		return string(append([]byte{x}, y...))
	}
	y := strconv.Itoa(1 + row)
	x := 'A' + byte(column)
	return string(append([]byte{x}, y...))
}

// OffsetY returns table y offset
func (t *Table) OffsetY() int {
	if t.Tag.OffsetY != 0 {
		return t.Tag.OffsetY
	}
	return 0
}

// OffsetX returns table x offset
func (t *Table) OffsetX() int {
	if t.Tag.OffsetY != 0 {
		return t.Tag.OffsetX
	}
	return 0
}

// NewTable creates a table
func NewTable(sliceType reflect.Type, tableTag *Tag, stylizer *Stylizer, parent *Column) (*Table, error) {
	structType := ensureStruct(sliceType)
	xStruct := xunsafe.NewStruct(structType)
	var ret = &Table{Tag: tableTag, Stylizer: stylizer, Type: sliceType, Parent: parent}
	ret.Columns = make(Columns, len(xStruct.Fields))
	for i := range xStruct.Fields {
		field := &xStruct.Fields[i]
		fieldTag, err := parseTag(field.Tag.Get(TagName))
		if err != nil {
			return nil, err
		}
		columnPos := int(field.Index)
		if pos := fieldTag.Position; pos != nil {
			columnPos = *pos
		}

		var xType *xunsafe.Type
		if field.Kind() == reflect.Ptr {
			xType = xunsafe.NewType(field.Type.Elem())
		}
		column := &Column{
			Field:    field,
			xType:    xType,
			Tag:      fieldTag,
			Position: columnPos,
		}
		ret.Columns[columnPos] = column
		if ensureSlice(field.Type) != nil && ensureStruct(field.Type) != nil {
			if err != nil {
				return nil, err
			}
			if column.Table, err = NewTable(field.Type, fieldTag, stylizer, column); err != nil {
				return nil, err
			}
			if column.Table.Tag.Name == "" || column.Table.Tag.Embed {
				column.Table.Tag.Embed = true
				column.Table.Tag.Name = ret.SheetName()
			}
			continue
		}

		if column.Tag.Style, err = stylizer.styleDefinition(column.Tag.Style, column.Tag.StyleRef); err != nil {
			return nil, err
		}
		column.setName(fieldTag, field, tableTag, parent)
		if style := column.Style(); style != nil {
			if err = stylizer.Register(style); err != nil {
				return nil, err
			}
		}
	}
	sort.Slice(ret.Columns, func(i, j int) bool {
		return ret.Columns[i].Position < ret.Columns[j].Position
	})
	return ret, nil
}

func (c *Column) setName(columnTag *Tag, field *xunsafe.Field, tableTag *Tag, parent *Column) {
	name := columnTag.Name
	if name == "" {
		name = field.Name
		if tableTag.Embed {
			name = parent.Field.Name + "." + name
		}
	}
	c.Name = name
}
