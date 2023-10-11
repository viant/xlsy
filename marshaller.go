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
	session []Option
}

// Marshal marshall arbitrary type to xls
func (m *Marshaller) Marshal(any interface{}) ([]byte, error) {
	dest := excelize.NewFile()
	rawType := reflect.TypeOf(any)
	stylizer := &Stylizer{registry: map[string]*Style{}, file: dest}
	aSession := newSession(nil, stylizer, NewTag())
	err := aSession.apply(m.session)
	if err != nil {
		return nil, err
	}
	if rawType.Kind() == reflect.Ptr {
		rawType = rawType.Elem()
	}
	var sheet *workSheet
	switch rawType.Kind() {
	case reflect.Slice:
		sheet, err = m.buildSheet(any, rawType, aSession)
	case reflect.Struct:
		sheet, err = m.buildSheets(any, rawType, aSession)
	default:
		return nil, fmt.Errorf("unsupported type: %T", any)
	}
	if err != nil {
		return nil, err
	}

	for _, name := range aSession.names {
		item := aSession.sheets[name]
		if err := item.transfer(); err != nil {
			return nil, err
		}
		if sheet == nil || sheet.index == nil {
			sheet = aSession.sheets[name]
		}
	}
	if sheet != nil {
		m.deleteDefaultWorksheetIfNeeded(sheet)
		sheet.SetActiveSheet()
	}

	err = dest.Close()
	buffer := new(bytes.Buffer)
	if err = dest.Write(buffer); err != nil {
		return nil, err
	}
	return buffer.Bytes(), err
}

func (m *Marshaller) deleteDefaultWorksheetIfNeeded(aSheet *workSheet) {
	if aSheet.index == nil {
		_ = aSheet.dest.DeleteSheet(defaultSheetName)
		return
	}
	workSheet := aSheet.dest.WorkBook.Sheets.Sheet[*aSheet.index]
	if workSheet.Name != defaultSheetName {
		_ = aSheet.dest.DeleteSheet(defaultSheetName)
	}
}

func (m *Marshaller) buildSheets(v any, structType reflect.Type, parent *session) (*workSheet, error) {
	xStruct := xunsafe.NewStruct(structType)
	ptr := xunsafe.AsPointer(v)
	var aSheet *workSheet
	for i := range xStruct.Fields {
		field := &xStruct.Fields[i]
		fieldType := field.Type
		if fieldType.Kind() == reflect.Ptr {
			fieldType = fieldType.Elem()
		}

		tag, err := parseTag(field.Tag)
		if err != nil {
			return nil, err
		}
		if tag.WorkSheet == "" {
			tag.WorkSheet = field.Name
		}
		value := field.Value(ptr)
		switch fieldType.Kind() {
		case reflect.Struct:
		case reflect.Slice:
		default:
			continue
		}
		aSession := newSession(parent, parent.stylizer, tag)
		ret, err := m.buildSheet(value, field.Type, aSession)
		if err != nil {
			return nil, err
		}
		if aSheet == nil {
			aSheet = ret
		}
		if err != nil {
			return nil, err
		}
	}
	if aSheet != nil {
		return aSheet, nil
	}
	return aSheet, nil
}

func (m *Marshaller) buildSheet(v any, tableType reflect.Type, aSession *session) (*workSheet, error) {
	aTable, err := NewTable(tableType, aSession.tag, aSession, nil)
	if err != nil {
		return nil, err
	}
	if aTable.Omitempty && v == nil {
		return nil, err
	}
	sheetName := aTable.SheetName()
	aSheet, err := aSession.getOrCreateSheet(sheetName, aTable.First)
	if err != nil {
		return nil, err
	}
	aSheet.addTable(aTable)
	if err = m.setTableHeader(aTable, aSession); err != nil {
		return nil, err
	}
	if err = m.setTableData(v, aTable); err != nil {
		return nil, err
	}

	return aSheet, nil
}

func (m *Marshaller) setTableData(v any, aTable *Table) error {
	if aTable.IsStruct {
		if v == nil {
			return nil
		}
		ptr := xunsafe.AsPointer(v)
		if ptr == nil {
			return nil
		}
		return m.setRecord(aTable, 0, ptr)
	}

	sliceType := ensureSlice(aTable.Type)

	xSlice := xunsafe.NewSlice(sliceType)
	ptr := xunsafe.AsPointer(v)
	sliceLen := xSlice.Len(ptr)
	aTable.Cardinality = sliceLen
	for sliceIndex := 0; sliceIndex < sliceLen; sliceIndex++ {
		record := xSlice.ValueAt(ptr, sliceIndex)
		if record == nil {
			continue
		}
		recordPtr := xunsafe.AsPointer(record)
		if recordPtr == nil {
			continue
		}
		err := m.setRecord(aTable, sliceIndex, recordPtr)
		if err != nil {
			return err
		}
	}
	return nil
}

func (m *Marshaller) setRecord(aTable *Table, sliceIndex int, recordPtr unsafe.Pointer) error {
	columnOffset := 0
	for i := 0; i < len(aTable.Columns); i++ {
		column := aTable.Columns[i]
		if column.Tag.Ignore {
			continue
		}
		if column.Tag.Blank {
			columnOffset++
			continue
		}
		row := aTable.Rows.index(sliceIndex)
		xField := column.Field
		value := xField.Value(recordPtr)
		if column.Table != nil {
			if err := m.setTableData(value, column.Table); err != nil {
				return err
			}
			if !column.Table.IsStandalone() {
				cell := row.Values.index(columnOffset)
				cell.rows = column.Table.Rows
				columnOffset++
			}
			continue
		}
		cell := row.Values.index(columnOffset)
		cell.omitEmpty = column.Tag.Omitempty
		columnOffset++
		if column.Table != nil {
			cell.rows = column.Table.Rows
			column.Table.Rows = nil
			continue
		}
		isNil := false
		if xField.Kind() == reflect.Ptr {
			if (*unsafe.Pointer)(xunsafe.AsPointer(value)) != nil {
				value = column.xType.Deref(value)
			} else {
				isNil = true
			}
		}
		if !isNil {
			cell.setValue(value)
		}
		if styleID := column.CellStyleID(aTable.Stylizer); styleID != nil {
			cell.styleID = styleID
		}
	}
	return nil
}

func (m *Marshaller) setTableHeader(aTable *Table, aSession *session) error {
	columnOffset := 0
	for i := 0; i < len(aTable.Columns); i++ {
		column := aTable.Columns[i]
		column.Position = i
		if column.Tag.Ignore {
			continue
		}
		header := aTable.newHeader(columnOffset, i)
		columnOffset++
		if column.Tag.Blank {
			continue
		}
		if column.Table != nil {
			m.setHeader(column.Table, header, column)

			if err := m.setTableHeader(column.Table, aSession); err != nil {
				return err
			}
			if !column.Table.IsStandalone() {
				header.header = column.Table.Header
			} else {
				sheetName := column.Table.SheetName()
				aSheet, err := aSession.getOrCreateSheet(sheetName, column.Table.First)
				if err != nil {
					return err
				}
				aSheet.addTable(column.Table)
			}
			continue
		}

		m.setHeader(aTable, header, column)
	}
	return nil
}

func (m *Marshaller) setHeader(aTable *Table, header *value, column *Column) {
	header.value = column.Name
	if styleID := column.HeaderStyleID(aTable.Stylizer); styleID != nil {
		header.styleID = styleID
	}
	if width := column.Width(aTable.Stylizer); width != nil {
		header.width = width.Value()
	}
}

// NewMarshaller create a marshaler with option
func NewMarshaller(opts ...Option) *Marshaller {
	opts = append(opts, WithDefaultHeaderStyle("font-style:bold"))
	ret := &Marshaller{session: opts}
	return ret

}
