package main

import (
	"fmt"
	"os"
	"errors"
	"strings"

	"github.com/nari-z/gable/gable"
)

const (
	cmdExport = "export"
	cmdImport = "import"

	usageError = "Usage: gable (export|import) SourceFilePath DestFilePath"
	notFoundFileError = "not found: %s";
	mismatchFileFormatError = "file format mismatch: %s src=%s dst=%s";
)

type execArgs struct {
	ExecCmd string
	SrcFilePath string
	DstFilePath string
}

func main() {
	arg, err := getExecArgs();

	if err != nil {
		fmt.Println("args error.");
		fmt.Println(err.Error());
		os.Exit(1);
	}

	switch arg.ExecCmd {
	case cmdExport: err = gable.Export(arg.SrcFilePath, arg.DstFilePath);
	case cmdImport: err = gable.Import(arg.SrcFilePath, arg.DstFilePath);
	}

	if err != nil {
		fmt.Println("exec error.");
		fmt.Println(err.Error());
		os.Exit(1);
	}

	fmt.Println("Done.");
}

func getExecArgs() (*execArgs, error) {
	if len(os.Args) != 4 {
		return nil, errors.New(usageError);
	}

	result := &execArgs{
		ExecCmd: os.Args[1],
		SrcFilePath: os.Args[2],
		DstFilePath: os.Args[3],
	};

	if (result.ExecCmd != cmdExport) && (result.ExecCmd != cmdImport) {
		return nil, errors.New(usageError);
	}
	_, err := os.Stat(result.SrcFilePath);
	if os.IsNotExist(err) == true {
		return nil, errors.New(fmt.Sprintf(notFoundFileError, result.SrcFilePath));
	}

	switch result.ExecCmd {
	case cmdExport:
		if (isExcelFile(result.SrcFilePath) == false) || (isXMLFile(result.DstFilePath) == false) {
			return nil, errors.New(fmt.Sprintf(mismatchFileFormatError,
											   result.ExecCmd,
											   "ExcelFile",
											   "XMLFile"));
		}
	case cmdImport:
		if (isXMLFile(result.SrcFilePath) == false) || (isExcelFile(result.DstFilePath) == false) {
			return nil, errors.New(fmt.Sprintf(mismatchFileFormatError,
											   result.ExecCmd,
											   "XMLFile",
											   "ExcelFile"));
		}
	}

	return result, nil;
}

func isExcelFile(filePath string) bool {
	targetExt := filePath[strings.LastIndex(filePath, "."):];
	exts := []string{".xlsx", ".xlsm"};
	for _, ext := range exts {
		if targetExt == ext {
			return true;
		}
	}

	return false;
}

func isXMLFile(filePath string) bool {
	targetExt := filePath[strings.LastIndex(filePath, "."):];
	return targetExt == ".xml";
}

