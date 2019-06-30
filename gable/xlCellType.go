package gable

type XlCellType struct {
	Value string `xml:",chardata"`
}

// TODO: 列挙している値のみ使用するよう制限。enumなどを使用する？
var TypeString        = XlCellType{"String"}
var TypeFormula       = XlCellType{"Formula"}
var TypeStringFormula = XlCellType{"StringFormula"}
var TypeNumeric       = XlCellType{"Numeric"}
var TypeBool          = XlCellType{"Bool"}
var TypeInline        = XlCellType{"Inline"}
var TypeError         = XlCellType{"Error"}
var TypeDate          = XlCellType{"Date"}

func (c XlCellType) ToString() string {
    if c.Value == "" {
        return "";
	}
	
    return c.Value;
}

func (c XlCellType) Equal(target XlCellType) bool {
	return c.Value == target.Value;
}