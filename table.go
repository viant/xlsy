package xlsy

import (
	"github.com/viant/xreflect"
	"github.com/viant/xunsafe"
	"reflect"
	"sort"
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
		indexPos indexPos
		Parent   *Column
		Stylizer *Stylizer
		*Tag
		Columns     Columns
		Header      *Row
		Rows        Rows
		Type        reflect.Type
		IsStruct    bool
		Cardinality int
	}

	indexPos []int
	Rows     []*Row

	Values []*value
	Row    struct {
		Values   Values
		Snapshot Cursor
		StyleID  int
	}

	value struct {
		header    *Row
		rows      Rows
		hasValue  bool
		value     interface{}
		styleID   *int
		width     float64
		omitEmpty bool
		snapshot  *Cursor
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

func (v *value) Columns() int {
	if v.header != nil {
		count := 0
		for _, candidate := range v.header.Values {
			if !candidate.omitEmpty {
				count++
			}
			if candidate.hasValue {
				count++
			}
		}
		return count
	}
	return 1
}

func (v *value) Rows() int {
	if rows := len(v.rows); rows > 0 {
		return rows
	}
	return 1
}

func (v *value) setValue(val interface{}) {
	v.value = val
	v.hasValue = true
}

func (v *value) HasValue() bool {
	if v.hasValue {
		return true
	}
	if len(v.rows) == 0 {
		return false
	}
	if len(v.rows) > 1 {
		return true
	}
	for _, value := range v.rows[0].Values {
		if value.HasValue() {
			return true
		}
	}
	return false
}

func (v *Values) index(index int) *value {
	if index < len(*v) {
		return (*v)[index]
	}
	for i := len(*v); i < index+1; i++ {
		*v = append(*v, &value{})
	}
	return (*v)[index]
}

func (v *Rows) index(index int) *Row {
	if index < len(*v) {
		return (*v)[index]
	}
	for i := len(*v); i < index+1; i++ {
		*v = append(*v, &Row{})
	}
	return (*v)[index]
}

func (v *indexPos) expand(index int) {
	if index < len(*v) {
		return
	}
	for i := len(*v); i < index+1; i++ {
		*v = append(*v, 0)
	}
}

// HeaderStyleID returns header style ID or nil
func (c *Column) HeaderStyleID(stylizer *Stylizer) *int {
	if c.Tag.HeaderStyle == nil {
		return nil
	}
	if style := c.Tag.HeaderStyle.Style; style != "" {
		if style := stylizer.Style(style); style != nil && style.Header.Style != nil {
			return style.Header.ID
		}
	}
	return nil
}

// ColumnStyleID returns column style ID or nil
func (c *Column) ColumnStyleID(stylizer *Stylizer) *int {
	if c.Tag.ColumnStyle == nil {
		return nil
	}
	if style := c.Tag.ColumnStyle.Style; style != "" {
		if style := stylizer.Style(style); style != nil && style.Column.Style != nil {
			return style.Column.ID
		}
	}
	return nil
}

// Width returns width
func (c *Column) Width(stylizer *Stylizer) *Length {
	style := c.Tag.ColumnStyle
	if style == nil {
		return nil
	}
	if style.Style != "" {
		if style := stylizer.Style(style.Style); style != nil && style.Column != nil {
			width := style.Column.Width
			if style.Column.WidthMax != nil {
				if style.Column.WidthMax.Value() < width.Value() {
					return style.Column.WidthMax
				}
			}
			return style.Column.Width
		}
	}
	return nil
}

// CellStyleID returns a style ID
func (c *Column) CellStyleID(stylizer *Stylizer) *int {
	if c.Tag.CellStyle == nil {
		return nil
	}
	if style := c.Tag.CellStyle.Style; style != "" {
		if style := stylizer.Style(style); style != nil && style.Cell.Style != nil {
			return style.Header.ID
		}
	}

	return nil
}

// SheetName returns table cheed name
func (t *Table) SheetName() string {
	if t.Tag.WorkSheet != "" {
		return t.Tag.WorkSheet
	}
	return defaultSheetName
}

func (t *Table) IsStandalone() bool {
	return t.Tag.WorkSheet != ""
}

// OffsetY returns table y dim
func (t *Table) OffsetY() int {
	if t.Tag.RowOffset != 0 {
		return t.Tag.RowOffset
	}
	return 0
}

// OffsetX returns table x dim
func (t *Table) OffsetX() int {
	if t.Tag.RowOffset != 0 {
		return t.Tag.ColumnOffset
	}
	return 0
}

func (t *Table) newHeader(index, position int) *value {
	if t.Header == nil {
		t.Header = &Row{}
	}
	value := t.Header.Values.index(index)
	t.indexPos.expand(position)
	t.indexPos[position] = index
	return value
}

func (t *Table) UseRow(b bool) bool {
	if t.Invert() {
		return !b
	}
	return b
}

// NewTable creates a table
func NewTable(rType reflect.Type, tableTag *Tag, aSession *session, parent *Column) (*Table, error) {

	isTime := xreflect.TimeType == rType || xreflect.TimePtrType == rType
	isStruct := !isTime && ensureStruct(rType) != nil && rType.Kind() != reflect.Slice
	structType := ensureStruct(rType)
	xStruct := xunsafe.NewStruct(structType)
	var ret = &Table{Tag: tableTag, Stylizer: aSession.stylizer, Type: rType, Parent: parent, IsStruct: isStruct}
	ret.Columns = make(Columns, len(xStruct.Fields))
	for i := range xStruct.Fields {
		field := &xStruct.Fields[i]
		fieldTag, err := parseTag(field.Tag)
		if err != nil {
			return nil, err
		}

		if fieldTag.Inverted == nil {
			fieldTag.Inverted = tableTag.Inverted
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

		isTime := xreflect.TimeType == field.Type || xreflect.TimePtrType == field.Type
		isStruct := !isTime && ensureStruct(field.Type) != nil
		column.setName(fieldTag, field)
		if isStruct && (!fieldTag.Ignore && !fieldTag.Blank) {
			if column.Table, err = NewTable(field.Type, fieldTag, aSession, column); err != nil {
				return nil, err
			}
			continue
		}

		if style := column.Tag.CellStyle; style != nil {
			if style.Style, err = aSession.stylizer.styleDefinition(style.Destination, style.Style, style.Ref); err != nil {
				return nil, err
			}
		}
		if column.Tag.HeaderStyle == nil {
			column.Tag.HeaderStyle = &StyleTag{Destination: "header"}
		}
		if style := column.Tag.HeaderStyle; style != nil {
			if style.Style, err = aSession.stylizer.styleDefinition(style.Destination, style.Style, style.Ref); err != nil {
				return nil, err
			}
		}
		if style := column.Tag.CellStyle; style != nil {
			if err = aSession.stylizer.Register(style.Definition()); err != nil {
				return nil, err
			}
		}
		if style := column.Tag.HeaderStyle; style != nil {
			if err = aSession.stylizer.Register(style.Definition()); err != nil {
				return nil, err
			}
		}
		if style := column.Tag.ColumnStyle; style != nil {
			if err = aSession.stylizer.Register(style.Definition()); err != nil {
				return nil, err
			}
		}
	}
	sort.Slice(ret.Columns, func(i, j int) bool {
		return ret.Columns[i].Position < ret.Columns[j].Position
	})
	return ret, nil
}

func (c *Column) setName(columnTag *Tag, field *xunsafe.Field) {
	if columnTag.Inline {
		return
	}
	name := columnTag.FormatName()
	if name == "" {
		name = field.Name
	}
	c.Name = name
}
