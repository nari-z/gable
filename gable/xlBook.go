package gable

import (
	"encoding/xml"
)

type XlBook struct {
	XMLName xml.Name `xml:"Book"`
	Sheets []*XlSheet `xml:"Sheets>Sheet"`
}

func (b *XlBook) AddSheet(newSheet *XlSheet) {
	b.Sheets = append(b.Sheets, newSheet);
}
