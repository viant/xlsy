package xlsy

import (
	"github.com/viant/parsly"
	"github.com/viant/parsly/matcher"
)

const (
	whitespaceToken = iota
	comaTerminatorToken
	semicolonTerminatorToken
	colonTerminatorToken
	equalTerminatorToken
	scopeBlockToken
	singleQuotedToken
	numberToken
)

var (
	numberMatcher              = parsly.NewToken(numberToken, "Number", matcher.NewNumber())
	whitespaceMatcher          = parsly.NewToken(whitespaceToken, "Whitespace", matcher.NewWhiteSpace())
	comaTerminatorMatcher      = parsly.NewToken(comaTerminatorToken, "coma", matcher.NewTerminator(',', true))
	equalTerminatorMatcher     = parsly.NewToken(equalTerminatorToken, "=", matcher.NewTerminator('=', true))
	colonTerminatorMatcher     = parsly.NewToken(colonTerminatorToken, "colon", matcher.NewTerminator(':', true))
	semicolonTerminatorMatcher = parsly.NewToken(semicolonTerminatorToken, "semicolon", matcher.NewTerminator(';', true))
	scopeBlockMatcher          = parsly.NewToken(scopeBlockToken, "{ .... }", matcher.NewBlock('{', '}', '\\'))
	singleQuotedMatcher        = parsly.NewToken(singleQuotedToken, "single quoted block", matcher.NewQuote('\'', '\''))
)
