package xlsy

import (
	"bytes"
	"context"
	"fmt"
	"github.com/stretchr/testify/assert"
	"github.com/viant/afs"
	"github.com/viant/afs/file"
	"github.com/viant/structology/format"
	"os"
	"path"
	"testing"
	"time"
)

func TestMarshaller_Marshal(t *testing.T) {

	now := time.Now().Truncate(time.Hour)
	var testCases = []struct {
		description string
		get         func() interface{}
		options     []Option
	}{

		{
			description: "slice",
			options: []Option{
				WithTag(&Tag{Tag: &format.Tag{Name: "Document"}}), WithInverted(),
			},
			get: func() interface{} {
				type Item struct {
					Sequence    int    `xls:"Seq"`
					Description string `xls:"-"`
					Product     string
				}
				type Record struct {
					ID     int
					Name   string
					Amount float64
					List   []*Item
					Info   string
				}

				return []*Record{
					{
						ID:     1,
						Name:   "Name 1",
						Amount: 3.2,
						List: []*Item{
							{
								Sequence: 1,
								Product:  "P1",
							},
							{
								Sequence: 3,
								Product:  "P2",
							},
						},
						Info: "info 1",
					},
					{
						ID:     2,
						Name:   "Name 2",
						Amount: 6.5,
						List: []*Item{
							{
								Sequence: 1,
								Product:  "P4",
							},
							{
								Sequence: 2,
								Product:  "P2",
							},
						},
						Info: "info 2",
					}}

			},
		},

		{
			description: "one to one",
			options: []Option{
				WithTag(&Tag{WorkSheet: "Document"}),
			},
			get: func() interface{} {

				type Doc struct {
					ID       *int
					Customer string
					Date     *time.Time `xls:"style={width:150px}"`
					Comments string
				}

				type Holder struct {
					Doc *Doc
				}

				return &Holder{Doc: &Doc{
					ID:       intPtr(101),
					Customer: "customer 1",
					Comments: "commnets",
				}}
			},
		},

		{
			description: "object marshaling",
			get: func() interface{} {

				type Filter struct {
					Name    string
					Include []string
					Exclude []string
				}

				type Info struct {
					Report     string
					ReportDate *time.Time
					b1         struct{} `xls:"blank"`
					From       string
					To         string
					Filters    []*Filter
				}

				type Record struct {
					ID     int
					Name   string
					Active bool
				}

				type Holder struct {
					Info    *Info     `xls:"invert=true" `
					Records []*Record `xls:"name=Records"`
				}

				return &Holder{
					Info: &Info{
						Report:     "Total",
						ReportDate: &now,
						From:       "2023-08-01",
						To:         "2023-08-02",
						Filters: []*Filter{
							{
								Name:    "Col1",
								Include: []string{"1,2"},
							},
							{
								Name:    "Col2",
								Exclude: []string{"10,20"},
							},
						},
					},
					Records: []*Record{
						{
							ID:   1,
							Name: "name 1",
						},
						{
							ID:     2,
							Name:   "name 2",
							Active: true,
						},
					},
				}
			},
			options: []Option{WithInverted()},
		},

		{
			description: "slice marshaling",
			get: func() interface{} {
				type Foo struct {
					ID      int `xls:"name=Id"`
					Name    string
					Price   float64    `xls:"name=Price,style={width:200;color:red;font-style:bold;format:'###,##0.0000'}"`
					Started *time.Time `xls:"name=Started,style={format:iso8601}"`
				}
				return []*Foo{
					{
						ID:      1,
						Name:    "name 1",
						Started: &now,
					},
					{
						ID:    2,
						Name:  "name 2",
						Price: 1231232312.4444,
					},
				}

			},

			options: []Option{},
		},
	}
	fs := afs.New()

	for i, testCase := range testCases {

		aMarshaller := NewMarshaller(testCase.options...)
		data, err := aMarshaller.Marshal(testCase.get())
		if !assert.Nil(t, err, testCase.description) {
			continue
		}
		assert.Truef(t, len(data) > 0, testCase.description)
		baseDir, _ := os.Getwd()
		err = fs.Upload(context.Background(), path.Join(baseDir, "testdata", fmt.Sprintf("test_%02d.xlsx", i)), file.DefaultFileOsMode, bytes.NewReader(data))
		assert.Nil(t, err, testCase.description)
	}
}

func intPtr(i int) *int {
	return &i
}
