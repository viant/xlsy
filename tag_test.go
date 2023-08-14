package xlsy

import (
	"github.com/stretchr/testify/assert"
	"testing"
)

func TestParseTag(t *testing.T) {

	var testCases = []struct {
		description string
		tag         string
		expect      *Tag
	}{

		{
			tag:    `name=ColumnX`,
			expect: &Tag{Name: "ColumnX"},
		},
		{
			tag:    `style={width:10;height:30}`,
			expect: &Tag{Style: "width:10;height:30"},
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
