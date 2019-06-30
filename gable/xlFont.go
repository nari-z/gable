package gable

type XlFont struct {
	Size int
	Name string
	Family int
	Charset int
	Color string
	Bold bool
	Italic bool
	Underline bool

	isEdited bool
}

func NewFont(isApply bool) XlFont {
	return XlFont{isEdited: isApply};
}

func (f XlFont) IsEdited() bool {
	return f.isEdited;
}
