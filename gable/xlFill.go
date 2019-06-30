package gable

type XlFill struct {
	PatternType string
	BgColor string
	FgColor string

	isEdited bool
}

func NewFill(isApply bool) XlFill {
	return XlFill{isEdited: isApply};
}

func (f XlFill) IsEdited() bool {
	return f.isEdited;
}
