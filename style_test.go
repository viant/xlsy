package xlsy

import (
	"github.com/stretchr/testify/assert"
	"testing"
)

func TestStyle_Init(t *testing.T) {
	var testCases = []struct {
		description string
		style       *Style
	}{
		{
			description: "",
			style:       &Style{Definition: "width:35.5;height:30;color:red"},
		},
	}

	for _, testCase := range testCases {
		err := testCase.style.Init()
		assert.Nil(t, err, testCase.description)
	}
}
