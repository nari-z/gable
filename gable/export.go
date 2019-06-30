package gable

import (
	// "fmt"
	"os"
	"errors"
	// "time"
	"encoding/xml"
	"github.com/tealeg/xlsx"
)

func Export(srcFilePath string, dstFilePath string) error {
	// read excel file.
	book, err := readExcel(srcFilePath);
	if err != nil {
		return err;
	}

	// write xml file.
	return outputXML(book, dstFilePath);
}

func readExcel(filePath string) (*XlBook, error) {
    srcFile, err := xlsx.OpenFile(filePath);
    if err != nil {
		return nil, err;
    }
	var book XlBook;
	book = XlBook{};

    for _, srcSheet := range srcFile.Sheets {
		sheet := &XlSheet{Name: srcSheet.Name};
		book.AddSheet(sheet);

        for rowIndex, srcRow := range srcSheet.Rows {
            for colIndex, srcCell := range srcRow.Cells {
				cell := &XlCell{};

				cell.Row = rowIndex;
				cell.Column = colIndex;

				cell.Value, cell.Type, err = getCellValueAndType(srcCell);
				if err != nil {
					return nil, err;
				}

				cell.Style, err = getXlStyle(srcCell.GetStyle());
				if err != nil {
					return nil, err;
				}

				if cell.IsEdited() == false {
					continue;
				}
			
				sheet.AddCell(cell);
			}
        }
	}

	return &book, nil;
}

func outputXML(book *XlBook, filePath string) error {
    file, err := os.Create(filePath);
    if err != nil {
		return err;
    }
	defer file.Close();
	
	buf, err := xml.MarshalIndent(book, "", "  "); // MEMO: ファイルサイズが大きくなりそうならインデントをなくす。
	if err != nil {
		return err;
	}

	xmlText := string(buf);
	file.Write(([]byte)(xmlText));
	
	return nil;
}

func getCellValueAndType(cell *xlsx.Cell) (string, XlCellType, error) {
	var targetType xlsx.CellType;
	targetType = cell.Type();

	if targetFormula := cell.Formula(); targetFormula != "" {
		if targetType == xlsx.CellTypeStringFormula {
			return targetFormula, TypeStringFormula, nil;
		}
		return targetFormula, TypeFormula, nil;
	}
	if cell.IsTime() == true {
		return cell.Value, TypeDate, nil;
	}

	switch targetType {
	case xlsx.CellTypeString:
		return cell.Value, TypeString, nil;
	case xlsx.CellTypeNumeric:
		return cell.Value, TypeNumeric, nil;
	case xlsx.CellTypeBool:
		return cell.Value, TypeBool, nil;
	case xlsx.CellTypeInline:
		return cell.Value, TypeInline, nil;
	case xlsx.CellTypeError:
		return cell.Value, TypeError, nil;
	// TODO: 日付セル対応、日付であることを判定する手段が不明。
	//       CellType、IsTime()は使用不可。Numericとして扱われている。
	//       formatも他と変わらない。
	//       日付が判断できれば、time->string, string->time変換で実現できると思われる。
	// case xlsx.CellTypeDate:
	}

	return "", XlCellType{}, errors.New("unknown CellType");
}

func getXlStyle(src *xlsx.Style) (XlStyle, error) {
	var dst XlStyle;

	dst.Border = getXlBorder(src.Border);
	dst.Font = getXlFont(src.Font, src.ApplyFont);
	dst.Fill = getXlFill(src.Fill, src.ApplyFill);

	return dst, nil;
}

func getXlBorder(src xlsx.Border) XlBorder {
	var dst XlBorder;

	dst.Left        = src.Left;
	dst.LeftColor   = src.LeftColor;
	dst.Right       = src.Right;
	dst.RightColor  = src.RightColor;
	dst.Top         = src.Top;
	dst.TopColor    = src.TopColor;
	dst.Bottom      = src.Bottom;
	dst.BottomColor = src.BottomColor;

	return dst;
}

func getXlFont(src xlsx.Font, isApply bool) XlFont {
	var dst XlFont;
	dst = NewFont(isApply);

	dst.Size = src.Size;
	dst.Name = src.Name;
	dst.Family = src.Family;
	dst.Charset = src.Charset;
	dst.Color = src.Color;
	dst.Bold = src.Bold;
	dst.Italic = src.Italic;
	dst.Underline = src.Underline;

	return dst;
}

func getXlFill(src xlsx.Fill, isApply bool) XlFill {
	var dst XlFill;
	dst = NewFill(isApply);

	dst.PatternType = src.PatternType;
	dst.BgColor = src.BgColor;
	dst.FgColor = src.FgColor;

	return dst;
}

