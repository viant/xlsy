# xsly - XLS marshaler 
[![GoReportCard](https://goreportcard.com/badge/github.com/viant/xsly)](https://goreportcard.com/report/github.com/viant/xsly)
[![GoDoc](https://godoc.org/github.com/viant/xsly?status.svg)](https://godoc.org/github.com/viant/xsly)

This library is compatible with Go 1.17+

Please refer to [`CHANGELOG.md`](CHANGELOG.md) if you encounter breaking changes.

- [Motivation](#motivation)
- [Usage](#usage)
- [Contribution](#contributing-to-godiff)
- [License](#license)

## Motivation

This project was create to marshal arbitrary go struct/slice into XSL format.

## Introduction

This project uses [excelize](github.com/xuri/excelize) project to convert complex go type into XLS.

You can use the following struct Tag to customize output
- Name: header or sheet name
- Style: CSS like style that are translated to excelize.Style: i.e: style:{color:red;font-style:bold}
- Style: StyleRef: style reference
- Ignore: ignore struct field
- Blank: blank row or column
- Position: optional field column position
- Embed: by default relation are moved to separate sheet, embed flatten relation in the same sheet
- Direction: vertical uses rows as column
- OffsetX: initial row offset
- OffsetY: initial column offset

The following style are currently supported
- color
- background-color
- background-gradient
- font-style
- font-family
- vertical-align
- text-align
- text-wrap
- text-indent
- width
- width-max
- height
- format

## Usage

```go

package mypkg

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
		ID      int        `xls:"name=Id"`
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


```


## License

The source code is made available under the terms of the Apache License, Version 2, as stated in the file `LICENSE`.

Individual files may be made available under their own specific license,
all compatible with Apache License, Version 2. Please see individual files for details.



## Contributing to xsly

xsly is an open source project and contributors are welcome!

See [TODO](TODO.md) list

## License

The source code is made available under the terms of the Apache License, Version 2, as stated in the file `LICENSE`.

Individual files may be made available under their own specific license,
all compatible with Apache License, Version 2. Please see individual files for details.


## Credits and Acknowledgements

**Library Author:** Adrian Witas

