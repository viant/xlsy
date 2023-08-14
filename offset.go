package xlsy

type offset struct {
	row    int
	column int
	cols   int
	rows   int
}

func newOffset(row, column int) *offset {
	ret := &offset{}
	ret.row = row
	ret.column = column
	return ret
}
