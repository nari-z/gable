package gable

type XlStyle struct {
	Border XlBorder
	Font XlFont
	Fill XlFill
	// Alignment XlAlignment
}

func (s XlStyle) IsEdited() bool {
	if s.Border.IsEdited() == true {
		return true;
	}
	if s.Font.IsEdited() == true {
		return true;
	}
	if s.Fill.IsEdited() == true {
		return true;
	}

	return false;
}