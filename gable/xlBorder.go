package gable

const (
	initialValue = "none"
)

type XlBorder struct {
	Left        string
	LeftColor   string
	Right       string
	RightColor  string
	Top         string
	TopColor    string
	Bottom      string
	BottomColor string
}

func (b XlBorder) IsEdited() bool {
	if (b.Left != "" ) && (b.Left != initialValue) {
		return true;
	}
	if (b.Right != "") && (b.Right != initialValue) {
		return true;
	}
	if (b.Top != "") && (b.Top != initialValue) {
		return true;
	}
	if (b.Bottom != "") && (b.Bottom != initialValue) {
		return true;
	}

	return false;
}