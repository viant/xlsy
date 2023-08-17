package xlsy

import (
	"bytes"
	"context"
	"fmt"
	"github.com/stretchr/testify/assert"
	"github.com/viant/afs"
	"github.com/viant/afs/file"
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
			description: "one to one",
			options: []Option{
				WithTag(&Tag{Name: "Documnt"}),
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
				}}
			},
		},
		{
			description: "one to many with style ref",
			options: []Option{
				WithNamedStyles("header", "header-font-style:bold"),
				WithTag(&Tag{Name: "Documnt"}),
			},
			get: func() interface{} {

				type LineItem struct {
					Pos      int     `xls:"styleref=header"`
					Name     string  `xls:"styleref=header"`
					Quantity int     `xls:"styleref=header"`
					Price    float64 `xls:"styleref=header"`
				}

				type Doc struct {
					ID       *int        `xls:"styleref=header"`
					Customer string      `xls:"styleref=header"`
					Date     *time.Time  `xls:"styleref=header,style={width:150px}"`
					Items    []*LineItem `xls:"embed=true,styleref=header"`
					Comments string      `xls:"styleref=header"`
				}

				return []*Doc{
					{
						ID:       intPtr(101),
						Customer: "customer 1",
						Items: []*LineItem{
							{
								Pos:      1,
								Name:     "Item 1",
								Quantity: 3,
								Price:    10.3,
							},
							{
								Pos:      2,
								Name:     "Item 2",
								Quantity: 4,
								Price:    40.1,
							},
						},
						Comments: "comments 1",
					},
					{
						ID:       intPtr(101),
						Customer: "customer 2",
						Date:     &now,
						Items: []*LineItem{
							{
								Pos:      10,
								Name:     "Item 10",
								Quantity: 24,
								Price:    44.3,
							},
							{
								Pos:      20,
								Name:     "Item 20",
								Quantity: 33,
								Price:    14.3,
							},
						},
						Comments: "comments 2",
					},
				}
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
					Info    []*Info   `xls:"dir=Vertical" `
					Records []*Record `xls:"name=Records"`
				}

				return &Holder{
					Info: []*Info{
						{
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
			options: []Option{WithTag(&Tag{Direction: DirectionVertical})},
		},

		{
			description: "slice marshaling",
			get: func() interface{} {
				type Foo struct {
					ID      int `xls:"name=Id"`
					Name    string
					Price   float64    `xls:"name=Price,style={width:200;color:red;header-font-style:bold;format:'###,##0.0000'}"`
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
		{
			description: " marshaling with default style",
			get: func() interface{} {
				type Foo struct {
					Id      int `xls:"name=ID"`
					Name    string
					Price   float64    `xls:"name=Price,style={width:20;color:red;header-font-style:bold;format:'###,##0.0000'}"`
					Started *time.Time `xls:"name=Started,style={width:100px;format:iso8601}"`
				}
				return []*Foo{
					{
						Id:      1,
						Name:    "name 1",
						Started: &now,
					},
					{
						Id:    2,
						Name:  "name 2",
						Price: 1231232312.4444,
					},
				}

			},

			options: []Option{
				WithDefaultStyle("header-font-style:bold"),
			},
		},
	}
	fs := afs.New()

	for i, testCase := range testCases[:1] {
		aMarshaller := NewMarshaller(testCase.options...)
		data, err := aMarshaller.Marshal(testCase.get())
		if !assert.Nil(t, err, testCase.description) {
			continue
		}
		assert.Truef(t, len(data) > 0, testCase.description)
		err = fs.Upload(context.Background(), path.Join(os.Getenv("HOME"), fmt.Sprintf("test_%02d.xlsx", i)), file.DefaultFileOsMode, bytes.NewReader(data))
		assert.Nil(t, err, testCase.description)
	}
}

func intPtr(i int) *int {
	return &i
}
