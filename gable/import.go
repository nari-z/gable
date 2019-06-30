package gable

import (
	// "fmt"
	"os"
	"strings"
	"strconv"

	"io/ioutil"
	"encoding/xml"

	"github.com/tealeg/xlsx"
)

func Import(srcFilePath string, dstFilePath string) error {
	// read xml file.
	book, err := readXML(srcFilePath);
	if err != nil {
		return err;
	}

	// write excel file.
	return outputExcel(book, dstFilePath);
}

func readXML(filePath string) (*XlBook, error) {
    srcFile, err := os.Open(filePath);
    if err != nil {
		return nil, err;
	}
	defer srcFile.Close();

	buf, err := ioutil.ReadAll(srcFile);
	if err != nil {
		return nil, err;
	}

	var book XlBook;
	xml.Unmarshal(buf, &book);

	return &book, nil;
}

func outputExcel(book *XlBook, filePath string) error {
    dstFile, err := getFile(filePath);
    if err != nil {
		return err;
	}

    for _, srcSheet := range book.Sheets {
		dstSheet, err := getSheet(srcSheet.Name, dstFile);
		if err != nil {
			return err;
		}

        for _, srcCell := range srcSheet.Cells {
			dstCell := dstSheet.Cell(srcCell.Row, srcCell.Column);

			err = setCellValue(srcCell, dstCell);
			if err != nil {
				return err;
			}

			cellStyle, err := getCellStyle(srcCell.Style);
			if err != nil {
				return err;
			}
			dstCell.SetStyle(&cellStyle);
        }
	}

	return dstFile.Save(filePath);
}

func getFile(filePath string) (*xlsx.File, error) {
	_, err := os.Stat(filePath);
	if os.IsNotExist(err) == true {
		// ファイルが存在しない場合は新規作成
		return xlsx.NewFile(), nil;
	}

	// 既存ファイルを開く
    return xlsx.OpenFile(filePath);
}

func getSheet(sheetName string, book *xlsx.File) (*xlsx.Sheet, error) {
	var targetSheet *xlsx.Sheet;
	// 対象のシートを用意
	if targetSheet, exists := book.Sheet[sheetName]; exists == true {
		return targetSheet, nil;
	}
	targetSheet, err := book.AddSheet(sheetName);
	if err == nil {
		return targetSheet, nil;
	}

	// error
	return nil, err;
}

func setCellValue(src *XlCell, dst *xlsx.Cell) error {
	// TODO: 値がない場合もTypeを設定するようにする。
	switch src.Type.ToString() {
	case TypeString.ToString():
		dst.SetString(src.Value);
		return nil;
	case TypeBool.ToString():
		val, err := strconv.ParseBool(src.Value);
		if err != nil {
			return nil;
		}
		dst.SetBool(val);
		return nil;
	case TypeNumeric.ToString():
		return setNumeric(src, dst);
	// TODO: inlineとerrorの扱いがわからないので要調査。不要かも。
	// case TypeInline.ToString():
	// 	return nil;
	// case TypeError.ToString():
	// 	return nil;
	case TypeStringFormula.ToString():
		dst.SetStringFormula(src.Value);
		return nil;
	case TypeFormula.ToString():
		dst.SetStringFormula(src.Value);
		return nil;
	// TODO: 日付対応
	//       time->string, string->time変換で実現できると思われる・。
	// case xlsx.CellTypeDate:
	}

	return nil; // TODO: error message.
}

func setNumeric(src *XlCell, dst *xlsx.Cell) error {
	if src.Value == "" {
		return nil;
	}

	if strings.Contains(src.Value, ".") == true {
		// 小数
		val, err := strconv.ParseFloat(src.Value, 64);
		if err != nil {
			return err;
		}
		dst.SetValue(val);
		return nil;
	}

	// 整数
	val, err := strconv.Atoi(src.Value);
	if err != nil {
		return err;
	}
	dst.SetValue(val);
	return nil;
}

func getCellStyle(src XlStyle) (xlsx.Style, error) {
	var dst xlsx.Style;

	dst.Border = getCellBorder(src.Border);
	dst.Font = getCellFont(src.Font);
	dst.Fill = getCellFill(src.Fill);

	return dst, nil;
}

func getCellBorder(src XlBorder) xlsx.Border {
	dst := xlsx.Border{};

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

func getCellFont(src XlFont) xlsx.Font {
	dst := xlsx.Font{};

	dst.Size = src.Size;
	dst.Name = src.Name
	dst.Family = src.Family;
	dst.Charset = src.Charset;
	dst.Color = src.Color;
	dst.Bold = src.Bold;
	dst.Italic = src.Italic;
	dst.Underline = src.Underline;

	return dst;
}

func getCellFill(src XlFill) xlsx.Fill {
	dst := xlsx.Fill{};

	dst.PatternType = src.PatternType;
	dst.BgColor = src.BgColor;
	dst.FgColor = src.FgColor;

	return dst;
}


