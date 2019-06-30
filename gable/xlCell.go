package gable

import (
	"encoding/xml"
)

type XlCell struct {
	XMLName xml.Name `xml:"Cell"`
	Value string `xml:"Value"`
	Row int `xml:"row,attr"`
	Column int `xml:"col,attr"`
	Type XlCellType `xml:"Type"`
	Style XlStyle `xml:"Style"`
}

func (c *XlCell) IsEdited() bool {
	if c.Value != "" {
		return true;
	}
	if c.Style.IsEdited() == true {
		return true;
	}

	return false;
}
