package xlsy

import (
	"github.com/xuri/excelize/v2"
)

type workSheet struct {
	index *int
	name  string

	tables []*Table
	dest   *excelize.File
}

func (s *workSheet) addTable(table *Table) {
	if table.Tag.WorkSheet == "" {
		table.Tag.WorkSheet = table.SheetName()
	}
	s.tables = append(s.tables, table)
}

func (s *workSheet) SetActiveSheet() {
	if s.index == nil {
		return
	}
	s.dest.SetActiveSheet(*s.index)
}

func (s *workSheet) SetCellValue(cell string, value interface{}) error {
	if err := s.ensureWorksheet(); err != nil {
		return err
	}
	return s.dest.SetCellValue(s.name, cell, value)
}

func (s *workSheet) SetCellStyle(hCell, vCell string, styleID int) error {
	if err := s.ensureWorksheet(); err != nil {
		return err
	}
	return s.dest.SetCellStyle(s.name, hCell, vCell, styleID)
}

func (s *workSheet) MergeCells(hCell, vCell string) error {
	if err := s.ensureWorksheet(); err != nil {
		return err
	}
	return s.dest.MergeCell(s.name, hCell, vCell)
}

func (s *workSheet) SetColWidth(startCol, endCol string, width float64) error {
	if err := s.ensureWorksheet(); err != nil {
		return err
	}
	return s.dest.SetColWidth(s.name, startCol, endCol, width)
}

func (s *workSheet) transfer() error {
	loc := Cursor(0)
	for _, table := range s.tables {
		if err := s.transferTable(table, &loc); err != nil {
			return err
		}
	}
	return nil
}

func (s *workSheet) transferTable(table *Table, addr *Cursor) (err error) {
	table.Tag.adjustAddress(addr)

	cursor := addr.clone()
	headerDim, err := s.transferHeader(table, cursor)
	if err != nil {
		return err
	}
	cursor.inc(headerDim.value(table.UseRow(true)), table.UseRow(true))

	if _, err = s.transferData(table, cursor); err != nil {
		return err
	}

	return nil
}

func (s *workSheet) transferData(table *Table, cursor Cursor) (dim Cursor, err error) {

	for i := 0; i < len(table.Rows); i++ {
		callAddr := cursor.clone()
		row := table.Rows[i]
		height := 1

		for j, _ := range row.Values {
			if j >= len(table.Header.Values) {
				continue
			}
			header := table.Header.Values[j]
			cell := row.Values[j]
			column := table.columnByIndex(j)

			if i == 0 {
				if header.width > 0 {
					if err = s.SetColWidth(callAddr.columnAddr(), callAddr.columnAddr(), header.width); err != nil {
						return 0, err
					}
				}
			}

			hasValue := table.Rows[i].Values[j].HasValue()
			if column.Tag.Omitempty && !hasValue {
				continue
			}

			cellAddr := callAddr.String()
			colTable := column.Table
			if colTable != nil {
				subTableAddr := callAddr.clone()
				if colTable.Inverted != table.Inverted {
					if !colTable.Invert() {
						subTableAddr.setColumn(header.snapshot.column() + 1)
						subTableAddr.setRow(header.snapshot.row() + 1)
					}
				}
				colTable.Rows = cell.rows
				_, err = s.transferData(colTable, subTableAddr)
				if err != nil {
					return 0, err
				}
			}

			if cell.hasValue {
				if err = s.SetCellValue(cellAddr, cell.value); err != nil {
					return 0, err
				}
			}
			if cell.styleID != nil {
				if err = s.SetCellStyle(cellAddr, cellAddr, *cell.styleID); err != nil {
					return 0, err
				}
			}
			callAddr.inc(header.Columns(), table.UseRow(false))
			if cell.Rows() > height {
				height = cell.Rows()
			}
		}
		cursor.inc(height, table.UseRow(true))
	}
	return dim, nil
}

func (t *Table) columnByIndex(index int) *Column {
	pos := t.indexPos[index]
	column := t.Columns[pos]
	return column
}

func (s *workSheet) transferHeader(table *Table, cur Cursor) (dim Cursor, err error) {
	height := 1
	if table.Header == nil {
		return
	}
	for i, header := range table.Header.Values {
		span := 1
		if len(table.Rows) == 0 || i >= len(table.Rows[0].Values) {
			continue
		}
		column := table.columnByIndex(i)
		if column.Tag.Omitempty && !table.Rows[0].Values[i].HasValue() {
			continue
		}
		//TODO check omit empty with vertical to skip it
		header.snapshot = cur.clone().ptr()
		cellAddr := cur.String()
		if header.value != nil {
			if err = s.SetCellValue(cellAddr, header.value); err != nil {
				return 0, err
			}
			if header.styleID != nil {
				if err = s.SetCellStyle(cellAddr, cellAddr, *header.styleID); err != nil {
					return 0, err
				}
			}
		}

		if colTable := column.Table; colTable != nil {
			cellAddr := cur.clone()

			if !table.Inline {
				cellAddr.inc(1, table.UseRow(true))
			}
			cellDim, err := s.transferHeader(colTable, cellAddr)
			if err != nil {
				return 0, err
			}

			expanded := cur.clone()
			if h := cellDim.value(colTable.UseRow(true)); h >= height && !table.Inline {
				height = h + 1
			}
			if s := cellDim.value(colTable.UseRow(false)); s > span {
				span = s
				expanded.inc(s-1, colTable.UseRow(false))
			}

			if table.Inverted == column.Table.Inverted && !table.Inline {
				if err := s.MergeCells(cur.String(), expanded.String()); err != nil {
					return 0, err
				}
			}

		}
		dim.inc(span, table.UseRow(false))
		cur.inc(span, table.UseRow(false))
	}

	if !table.Inline {
		if err = s.mergeHeaders(table, height); err != nil {
			return 0, err
		}
	}
	dim.inc(height, table.UseRow(true))
	return dim, nil
}

func (s *workSheet) mergeHeaders(table *Table, height int) (err error) {
	if height < 1 { //nothing to merge
		return nil
	}
	// merge height
	for i, header := range table.Header.Values {
		if header.snapshot == nil {
			continue
		}
		if column := table.columnByIndex(i); column.Table != nil {
			continue
		}
		expanded := header.snapshot.clone()
		expanded.inc(height-1, table.UseRow(true))
		if err = s.MergeCells(header.snapshot.String(), expanded.String()); err != nil {
			return err
		}
	}
	return nil
}
