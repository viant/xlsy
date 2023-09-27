package xlsy_test

import (
	"bytes"
	"context"
	"github.com/viant/afs"
	"github.com/viant/afs/file"
	"github.com/viant/xlsy"
	"log"
	"time"
)

func ExampleNewMarshaller() {

	now := time.Now().Truncate(time.Second)
	type Record struct {
		ID      int        `xls:"name=ID"`
		Name    string     `xls:"name=Name"`
		Active  bool       `xls:"name=Active,style={color:blue}"`
		Started *time.Time `xls:"name=Started,style={width:200px;format:date}"`
	}
	var records = []*Record{
		{
			ID:      1,
			Name:    "Item 1",
			Active:  true,
			Started: &now,
		},
		{
			ID:      2,
			Name:    "Item 2",
			Active:  false,
			Started: &now,
		},
		{
			ID:      3,
			Name:    "Item 3",
			Active:  true,
			Started: &now,
		},
	}

	marshaller := xlsy.NewMarshaller(xlsy.WithDefaultStyle("font-style:bold"))
	data, err := marshaller.Marshal(records)
	if err != nil {
		log.Fatal(err)
	}
	fs := afs.New()
	err = fs.Upload(context.Background(), "test.xlsx", file.DefaultFileOsMode, bytes.NewReader(data))
	if err != nil {
		log.Fatal(err)
	}

}
