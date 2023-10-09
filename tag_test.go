package xlsy

import (
	"github.com/stretchr/testify/assert"
	"github.com/viant/structology/format"
	"reflect"
	"testing"
)

func TestParseTag(t *testing.T) {

	var testCases = []struct {
		description string
		tag         reflect.StructTag
		expect      *Tag
	}{

		{
			tag:    `xls:"name=ColumnX""`,
			expect: &Tag{Tag: &format.Tag{Name: "ColumnX"}},
		},
		{
			tag:    `xls:"style={width:10;height:30}"`,
			expect: &Tag{CellStyle: &StyleTag{Style: "width:10;height:30"}, Tag: &format.Tag{Name: ""}},
		},
	}

	for _, testCase := range testCases {
		actual, err := parseTag(testCase.tag)
		if !assert.Nil(t, err, testCase.description) {
			continue
		}
		if !assert.EqualValues(t, testCase.expect, actual, testCase.description) {
			continue
		}
	}
}
