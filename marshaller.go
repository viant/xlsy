package xlsy

import (
	bytes "bytes"
	"fmt"
	"github.com/viant/xunsafe"
	"github.com/xuri/excelize/v2"
	"reflect"
	"unsafe"
)

// Marshaller represents xls marshalle
type Marshaller struct {
	options []Option
}

// Marshal marshall arbitrary type to xls
func (m *Marshaller) Marshal(any interface{}, moptions ...interface{}) ([]byte, error) {
	dest := excelize.NewFile()
	rawType := reflect.TypeOf(any)
	stylizer := &Stylizer{registry: map[string]*Style{}, file: dest}
	opts := &options{stylizer: stylizer}
	err := opts.apply(m.options)
	if err != nil {
		return nil, err
	}
	if rawType.Kind() == reflect.Ptr {
		rawType = rawType.Elem()
	}
	offset := &offset{row: opts.tag.OffsetY, column: opts.tag.OffsetX}
	var sheet int
	switch rawType.Kind() {
	case reflect.Slice:
		sheet, err = m.marshalSlice(any, rawType, opts, stylizer, dest, offset)
	case reflect.Struct:
		sheet, err = m.marshalHolder(any, rawType, stylizer, dest)
	default:
		return nil, fmt.Errorf("unsupported type: %T", any)
	}
	if err != nil {
		return nil, err
	}
	m.deleteDefaultWorksheetIfNeeded(dest, sheet)
	dest.SetActiveSheet(sheet)
	err = dest.Close()
	buffer := new(bytes.Buffer)
	if err = dest.Write(buffer); err != nil {
		return nil, err
	}
	return buffer.Bytes(), err
}

func (m *Marshaller) deleteDefaultWorksheetIfNeeded(dest *excelize.File, sheet int) {
	workSheet := dest.WorkBook.Sheets.Sheet[sheet]
	if workSheet.Name != defaultSheetName {
		_ = dest.DeleteSheet(defaultSheetName)
	}
}

func (m *Marshaller) marshalHolder(v any, structType reflect.Type, stylizer *Stylizer, dest *excelize.File) (int, error) {
	xStruct := xunsafe.NewStruct(structType)
	ptr := xunsafe.AsPointer(v)
	var sheet *int
	for i := range xStruct.Fields {
		field := &xStruct.Fields[i]
		fieldType := field.Type
		if fieldType.Kind() == reflect.Ptr {
			fieldType = fieldType.Elem()
		}
		tag, err := parseTag(field.Tag.Get(TagName))
		if err != nil {
			return 0, err
		}
		if tag.Name == "" {
			tag.Name = field.Name
		}
		sliceType := fieldType
		value := field.Value(ptr)
		switch fieldType.Kind() {
		case reflect.Struct:
			sliceType = reflect.SliceOf(field.Type)
			value = m.toSingleElementSlice(sliceType, value)
		case reflect.Slice:
		default:
			continue
		}

		opts := &options{stylizer: stylizer, tag: tag}
		offset := &offset{row: tag.OffsetY, column: tag.OffsetX}
		ret, err := m.marshalSlice(value, sliceType, opts, stylizer, dest, offset)
		if sheet == nil {
			sheet = &ret
		}
		if err != nil {
			return 0, err
		}
	}
	if sheet != nil {
		return *sheet, nil
	}
	return 0, nil
}

func (m *Marshaller) toSingleElementSlice(sliceType reflect.Type, value interface{}) interface{} {
	slice := reflect.MakeSlice(sliceType, 1, 1)
	slice.Index(0).Set(reflect.ValueOf(value))
	slicePtrValue := reflect.New(sliceType)
	slicePtrValue.Elem().Set(slice)
	value = slicePtrValue.Interface()
	return value
}

func (m *Marshaller) marshalSlice(v any, sliceType reflect.Type, options *options, stylizer *Stylizer, dest *excelize.File, offset *offset) (int, error) {
	aTable, err := NewTable(sliceType, options.tag, stylizer, nil)
	if err != nil {
		return 0, err
	}
	sheetName := aTable.SheetName()
	// Create a new sheet.
	sheet, err := dest.NewSheet(sheetName)
	if err != nil {
		return 0, err
	}
	//	value := xSlice.ValueAt(ptr, i)
	if err = m.marshalTableHeader(aTable, dest, offset); err != nil {
		return 0, err
	}

	offset.row++
	if err = m.marshalTableData(v, aTable, dest, offset); err != nil {
		return 0, err
	}
	return sheet, nil
}

func (m *Marshaller) marshalTableData(v any, aTable *Table, dest *excelize.File, offset *offset) error {
	sliceType := ensureSlice(aTable.Type)
	xSlice := xunsafe.NewSlice(sliceType)
	ptr := xunsafe.AsPointer(v)
	sliceLen := xSlice.Len(ptr)
	sheetName := aTable.SheetName()
	//Render relations

	rowSpans, err := m.marshalRelations(sliceLen, xSlice, ptr, aTable, offset, dest)
	if err != nil {
		return err
	}
	rowOffset := offset.row
	for sliceIndex := 0; sliceIndex < sliceLen; sliceIndex++ {
		record := xSlice.ValueAt(ptr, sliceIndex)
		recordPtr := xunsafe.AsPointer(record)
		columnOffset := offset.column
		for i := 0; i < len(aTable.Columns); i++ {
			column := aTable.Columns[i]
			if column.Tag.Ignore {
				continue
			}
			xField := column.Field
			address := aTable.Address(rowOffset, columnOffset)
			columnOffset += column.Size()
			if column.Table != nil || column.Tag.Blank {
				continue
			}
			value := xField.Value(recordPtr)

			if xField.Kind() == reflect.Ptr {
				if (*unsafe.Pointer)(xunsafe.AsPointer(value)) != nil {
					value = column.xType.Deref(value)
				}
			}

			if err = dest.SetCellValue(sheetName, address, value); err != nil {
				return err
			}
			if styleID := column.CellStyleID(aTable.Stylizer); styleID != nil {
				if err = dest.SetCellStyle(sheetName, address, address, *styleID); err != nil {
					return err
				}
			}
		}
		rowOffset += rowSpans[sliceIndex]

	}
	offset.cols = len(aTable.Columns)
	offset.rows = sliceLen
	return nil
}

func (m *Marshaller) marshalRelations(sliceLen int, xSlice *xunsafe.Slice, ptr unsafe.Pointer, aTable *Table, offset *offset, dest *excelize.File) ([]int, error) {
	var rowSpans []int
	rowOffset := offset.row
	for sliceIndex := 0; sliceIndex < sliceLen; sliceIndex++ {
		record := xSlice.ValueAt(ptr, sliceIndex)
		recordPtr := xunsafe.AsPointer(record)
		rowSpan := 1

		columnOffset := offset.column
		for _, column := range aTable.Columns {
			if column.Tag.Ignore {
				continue
			}
			if column.Table == nil || column.Tag.Blank {
				columnOffset += column.Size()
				continue
			}
			value := column.Field.Value(recordPtr)
			childOffset := newOffset(rowOffset, columnOffset)
			if !column.Table.IsVertical() && aTable.IsVertical() {
				childOffset = newOffset(columnOffset, rowOffset)
			}
			columnOffset += column.Size()
			if err := m.marshalTableData(value, column.Table, dest, childOffset); err != nil {
				return nil, err
			}
			if rowSpan < childOffset.rows {
				rowSpan = childOffset.rows
			}
		}
		rowOffset += rowSpan
		rowSpans = append(rowSpans, rowSpan)
	}
	return rowSpans, nil
}

func (m *Marshaller) marshalTableHeader(aTable *Table, dest *excelize.File, offset *offset) error {
	sheetName := aTable.SheetName()
	columnOffset := offset.column
	rowOffset := offset.row
	for i := 0; i < len(aTable.Columns); i++ {
		column := aTable.Columns[i]
		if column.Tag.Ignore {
			continue
		}
		if column.Table != nil {
			fieldOffset := newOffset(rowOffset, columnOffset)
			if !column.Table.IsVertical() && aTable.IsVertical() {
				fieldOffset = newOffset(columnOffset, rowOffset)
			}
			if err := m.marshalTableHeader(column.Table, dest, fieldOffset); err != nil {
				return err
			}
			columnOffset += column.Size()
			continue
		}
		address := aTable.Address(rowOffset, columnOffset)
		columnAddress := aTable.ColumnAddress(columnOffset)
		columnOffset += column.Size()
		if column.Tag.Blank {
			continue
		}
		if err := dest.SetCellValue(sheetName, address, column.Name); err != nil {
			return err
		}
		if styleID := column.HeaderStyleID(aTable.Stylizer); styleID != nil {
			if err := dest.SetCellStyle(sheetName, address, address, *styleID); err != nil {
				return err
			}
		}
		if width := column.Width(aTable.Stylizer); width != nil {
			if err := dest.SetColWidth(sheetName, columnAddress, columnAddress, width.Value()); err != nil {
				return err
			}
		}
	}
	return nil
}

// NewMarshaller create a marshaler with option
func NewMarshaller(opts ...Option) *Marshaller {
	ret := &Marshaller{options: opts}
	return ret

}
