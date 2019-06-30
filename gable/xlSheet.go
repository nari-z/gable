package gable

import (
	"encoding/xml"
)

type XlSheet struct {
	XMLName xml.Name `xml:"Sheet"`
	Name string
	// Rows []*XlRow
	// Cols []*XlColumn
	Cells []*XlCell `xml:"Cells>Cell"`
}

func (b *XlSheet) AddCell(newCell *XlCell) {
	b.Cells = append(b.Cells, newCell);
}
