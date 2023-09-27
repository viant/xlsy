package xlsy

import (
	"fmt"
	"strconv"
)

type Cursor int

func (c *Cursor) clone() Cursor {
	return *c
}

func (c Cursor) ptr() *Cursor {
	return &c
}

func (c *Cursor) value(row bool) int {
	if row {
		return c.row()
	}
	return c.column()
}

func numberToColumn(num int) string {
	num++
	columnName := ""
	for num > 0 {
		num-- // adjust for 0-indexed counting
		remainder := num % 26
		columnName = string([]byte{byte(remainder + 65)}) + columnName
		num /= 26
	}
	if columnName == "" {
		return "A"
	}
	return columnName
}

func (c Cursor) String() string {
	//if inverted {
	//	y := strconv.Itoa(1 + c.column())
	//	x := numberToColumn(c.row())
	//	return string(append([]byte(x), y...))
	//}
	y := strconv.Itoa(1 + c.row())
	x := numberToColumn(c.column())
	return string(append([]byte(x), y...))
}

func (c Cursor) columnAddr() string {
	return numberToColumn(c.column())
}

func (c Cursor) RawString() string {
	return fmt.Sprintf("R:%v, C:%v\n", c.row(), c.column())
}

func (c *Cursor) set(row, column int) {
	*c = Cursor(packInt32sToInt64(int32(row), int32(column)))
}

func (c *Cursor) incRow(delta int) {
	c.set(c.row()+delta, c.column())
}

func (c *Cursor) incColumn(delta int) {
	c.set(c.row(), c.column()+delta)
}

func (c *Cursor) dim(row bool) int {
	if row {
		return c.row()
	}
	return c.column()
}

func (c *Cursor) inc(delta int, row bool) {
	if row {
		c.incRow(delta)
		return
	}
	c.incColumn(delta)
}

func (c *Cursor) setRow(row int) {
	column := c.column()
	*c = Cursor(packInt32sToInt64(int32(row), int32(column)))
}

func (c *Cursor) setColumn(column int) {
	row := c.row()
	*c = Cursor(packInt32sToInt64(int32(row), int32(column)))
}

func (c *Cursor) diff(begin Cursor) Cursor {
	rows := c.row() - begin.row()
	cols := c.column() - begin.column()
	ret := Cursor(0)
	ret.setRow(rows)
	ret.setColumn(cols)
	return ret
}

func (c *Cursor) row() int {
	row, _ := unpackInt64ToInt32s(int64(*c))
	return int(row)
}
func (c *Cursor) column() int {
	_, col := unpackInt64ToInt32s(int64(*c))
	return int(col)
}
