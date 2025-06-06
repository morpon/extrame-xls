package yymmdd

import "time"

func Parse(value, format string) (time.Time, error) {
	_, tokens := Lexer(format)
	ds := parse(tokens)
	return ds.Parse(value)
}

// Parse creates a new parser with the recommended
// parameters.
func parse(tokens []LexToken) Formatter {
	p := &parser{
		tokens: tokens,
		pos:    -1,
	}
	p.initState = initialParserState
	return p.run()
}

// run starts the statemachine
func (p *parser) run() Formatter {
	var f Formatter
	for state := p.initState; state != nil; {
		state = state(p, &f)
	}
	return f
}

// parserState represents the state of the scanner
// as a function that returns the next state.
type parserState func(*parser, *Formatter) parserState

// nest returns what the next token AND
// advances p.pos.
func (p *parser) next() *LexToken {
	if p.pos >= len(p.tokens)-1 {
		return nil
	}
	p.pos += 1
	return &p.tokens[p.pos]
}

// the parser type
type parser struct {
	tokens []LexToken
	pos    int
	serial int

	initState parserState
}

// the starting state for parsing
func initialParserState(p *parser, f *Formatter) parserState {
	var t *LexToken
	for t = p.next(); t[0] != T_EOF; t = p.next() {
		var item ItemFormatter
		switch t[0] {
		case T_YEAR_MARK:
			item = new(YearFormatter)
		case T_MONTH_MARK:
			item = new(MonthFormatter)
		case T_DAY_MARK:
			item = new(DayFormatter)
		case T_HOUR_MARK:
			item = new(HourFormatter)
		case T_MINUTE_MARK:
			item = new(MinuteFormatter)
		case T_SECOND_MARK:
			item = new(SecondFormatter)
		case T_RAW_MARK:
			item = new(basicFormatter)
		}
		item.setOriginal(t[1])
		f.Items = append(f.Items, item)
	}
	if len(t[1]) > 0 {
		r := new(basicFormatter)
		r.origin = t[1]
		f.Items = append(f.Items, r)
	}
	return nil
}
